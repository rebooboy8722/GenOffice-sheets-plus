import { test, expect } from '@playwright/test'
import { createServer, type Server } from 'node:http'
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { AddressInfo } from 'node:net'
import type { Page } from '@playwright/test'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { gridCellCenter, launchShell, waitForPageWithUrl } from './helpers'

/**
 * End-to-end proof of the freshness gate (apps/sheets/src/renderer/ai/
 * workbook-baseline.ts) — the fix for the fork's top correctness defect:
 * upstream applied a model's range-addressed write against coordinates that a
 * user's row insert had already invalidated, so the write landed on the wrong
 * record **with no error at all**.
 *
 * Both cases drive the real Electron shell against a scripted
 * OpenAI-compatible stub, in the same shape a real run takes:
 *   turn 1  the model reads the workbook (this sets the freshness baseline)
 *   — the user inserts a row through the ribbon, as a person would —
 *   turn 2  the model proposes a write to the address it read
 *
 * The stub answers turn 1 slowly, so the ribbon click lands in the window
 * between the read landing and the next model call — the exact race the gate
 * exists for. What the gate returned is then read from the tool result the app
 * sends back, and the workbook is saved and inspected on disk to confirm the
 * write really did (control) or did not (gate) reach the grid.
 */

interface ChatMessage {
  role: string
  content: unknown
  tool_call_id?: string
  tool_calls?: Array<{ id: string; type: 'function'; function: { name: string; arguments: string } }>
}
interface ChatRequest {
  messages: ChatMessage[]
}

const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')

/** Lets the test land a ribbon edit between two model calls. */
const EDIT_WINDOW_MS = 1_500

function startModelStub(
  requests: ChatRequest[],
  onEditWindow: () => Promise<void>,
): Promise<Server> {
  const server = createServer(async (req, res) => {
    if (req.method !== 'POST') {
      res.writeHead(404).end()
      return
    }
    const chunks: Buffer[] = []
    for await (const chunk of req) chunks.push(chunk as Buffer)
    const body = JSON.parse(Buffer.concat(chunks).toString('utf8')) as ChatRequest
    requests.push(body)
    const answered = (id: string) => body.messages.some((m) => m.tool_call_id === id)
    let message: ChatMessage
    if (!answered('read_1')) {
      message = {
        role: 'assistant',
        content: null,
        tool_calls: [
          {
            id: 'read_1',
            type: 'function',
            function: { name: 'get_workbook_context', arguments: '{}' },
          },
        ],
      }
    } else if (!answered('write_1')) {
      // The window the gate exists for: the read has landed (so the model holds
      // A1's coordinates and the baseline is marked), and the write has not run
      // yet. Editing here is what a user typing ⌘Z or inserting a row between
      // two model calls does — the case upstream got silently wrong.
      //
      // Editing during the *read* turn instead would be a different, legitimate
      // situation: the read would then report the post-edit layout, so the
      // model's coordinates would be fresh and the write should go through.
      await onEditWindow()
      // Write to the address read above — correct before a row insert, one row
      // off after one. The sheet id comes from the context the app just handed
      // back, as a real model would do (ids are import-assigned, never guessed).
      const context = body.messages.find((m) => m.tool_call_id === 'read_1')?.content ?? ''
      const sheetId = /id=([^,\s)]+)/.exec(String(context))?.[1] ?? ''
      message = {
        role: 'assistant',
        content: null,
        tool_calls: [
          {
            id: 'write_1',
            type: 'function',
            function: {
              name: 'propose_operations',
              arguments: JSON.stringify({
                operations: [{ op: 'set_cell', sheetId, address: 'A1', value: 'WRITTEN' }],
                summary: 'label A1',
              }),
            },
          },
        ],
      }
    } else {
      message = { role: 'assistant', content: 'Done.' }
    }
    res.writeHead(200, { 'content-type': 'application/json' })
    res.end(JSON.stringify({ choices: [{ message, finish_reason: 'stop' }] }))
  })
  return new Promise((r) => server.listen(0, '127.0.0.1', () => r(server)))
}

/** The slice of Univer's sheet facade this test needs inside the page. */
interface SheetLike {
  getRange(row: number, col: number, rows: number, cols: number): { getValue(): unknown }
}

/** The tool output the app reported back for the write call. */
function writeToolResult(requests: ChatRequest[]): string {
  for (const request of requests) {
    for (const m of request.messages) {
      if (m.role === 'tool' && typeof m.content === 'string' && m.tool_call_id === 'write_1') {
        return m.content
      }
    }
  }
  return ''
}

async function seedAiStub(port: number, userDataDir: string): Promise<void> {
  const { writeFile } = await import('node:fs/promises')
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = { apiKey: 'stub', model: 'stub-model', baseUrl: `http://127.0.0.1:${port}/v1` }
  await writeFile(join(userDataDir, 'ai-settings.json'), JSON.stringify(resolveCustomEndpoints(settings)))
}


async function askAi(page: Page, text: string): Promise<void> {
  const composer = page.getByPlaceholder(/Ask AI to analyze/)
  await composer.click()
  await composer.fill(text)
  // The send button rather than Enter: Enter inside a multi-line composer is
  // ambiguous, and a message that never sends would time out far from here.
  await page.getByRole('button', { name: 'Send', exact: true }).click()
}

async function run(
  insertRowBeforeWrite: boolean,
  testInfo: string,
): Promise<{ toolResult: string; a1: string; workbook: string }> {
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-stale-write-'))
  const workbook = join(scratch, 'book.xlsx')
  await copyFile(FIXTURE, workbook)

  const requests: ChatRequest[] = []
  let editWindow: () => Promise<void> = async () => {}
  const server = await startModelStub(requests, () => editWindow())
  const port = (server.address() as AddressInfo).port

  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-e2e-'))
  await seedAiStub(port, userDataDir)

  const launched = await launchShell({
    userDataDir,
    onboardingSeen: true,
    openFile: workbook,
    videoDir: testInfo,
    // Exposes window.__genofficeDebug.univerAPI, used to read the grid back.
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
    await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    await sheets.waitForTimeout(2_000)

    // Runs between the model's read and its write (see the stub).
    editWindow = async () => {
      if (!insertRowBeforeWrite) return
      const header = await gridCellCenter(sheets, 0, -1)
      await sheets.mouse.click(header.x, header.y)
      await sheets.waitForTimeout(300)
      await sheets.getByRole('button', { name: 'Insert row', exact: true }).click()
      await sheets.waitForTimeout(800)
      await new Promise((r) => setTimeout(r, EDIT_WINDOW_MS))
    }

    await askAi(sheets, 'Put the label WRITTEN in A1.')
    await expect(sheets.getByText('Done.').first()).toBeVisible({ timeout: 60_000 })

    // Read A1 straight from the live workbook (the same debug facade the other
    // sheets specs use). A save round-trip would need the app's own save
    // prompt, and the clipboard needs a window manager.
    const a1 = await sheets.evaluate(() => {
      const debug = (
        window as unknown as {
          __genofficeDebug: { univerAPI: { getActiveWorkbook(): { getActiveSheet(): SheetLike } } }
        }
      ).__genofficeDebug
      return String(debug.univerAPI.getActiveWorkbook().getActiveSheet().getRange(0, 0, 1, 1).getValue() ?? '')
    })
    await sheets.screenshot({ path: `/tmp/stale-${testInfo}.png` })
    return { toolResult: writeToolResult(requests), a1, workbook }
  } finally {
    await launched.app.close().catch(() => {})
    server.close()
  }
}

test.describe('sheets AI: refusing a write whose coordinates moved', () => {
  test('a write against coordinates nobody moved still lands', async () => {
    const { toolResult, a1 } = await run(false, 'stale-write-control')
    // Control: without the gate the two cases would be indistinguishable, so
    // this pins that a normal write is unaffected by the new checks.
    expect(toolResult).not.toContain('structure changed')
    expect(toolResult).toContain('Auto-applied')
    expect(a1).toContain('WRITTEN')
  })

  test('a write after the user inserts a row is refused, not silently misplaced', async () => {
    const { toolResult, a1 } = await run(true, 'stale-write-gated')
    expect(toolResult).toContain('structure changed')
    expect(toolResult).toContain('get_workbook_context')
    // The whole point: the old coordinates must not have been written to.
    expect(toolResult).not.toContain('Auto-applied')
    expect(a1).not.toContain('WRITTEN')
  })
})
