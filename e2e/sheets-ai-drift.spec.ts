import { expect, test } from '@playwright/test'
import { createServer, type Server } from 'node:http'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { AddressInfo } from 'node:net'
import type { Page } from '@playwright/test'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { buildGridWorkbook, gridCellCenter, launchShell, waitForPageWithUrl } from './helpers'

/**
 * The other two freshness scenarios (the plan's 验收 2 and 3), in the real app.
 *
 * Scenario 1 (`sheets-ai-stale-write.spec.ts`) covers an inserted row. These two
 * cover the cases that look like they should behave the same and do not:
 *
 *  - **a sort** moves every value in a range to a different row. No address
 *    moved, yet the model's belief "row 3 is alpha" is destroyed — and bulk ops
 *    like `fill_range` are exactly the ones whose meaning depends on the block's
 *    contents. Univer routes a sort through `sheet.mutation.reorder-range`,
 *    which the baseline counts as structural, so the write is refused.
 *  - **⌘Z after an AI write** leaves the addresses valid and the model's numbers
 *    wrong. That is content drift: the write stands, and the tool output says
 *    the ground moved.
 *
 * A scripted OpenAI-compatible stub makes the model's turns exactly what the
 * test says, and the user's edit lands *between* turns — the window the baseline
 * exists for. What the gate decided is read back from the tool output the app
 * sends to the model, which is the only channel the model itself can see.
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

/** What a scripted turn may need from the running app: the page (to land the
 * user's edit) and the sheet id. The id is assigned by the app on import, so a
 * test cannot hardcode it — a real model reads it out of the context too, and
 * the stub picks it up from the tool output the same way. */
interface TurnContext {
  page: Page
  sheetId: string
}

/** One scripted model turn. It matches once `after` has been answered and
 * `until` has not; `edit` runs just before answering, i.e. between turns. */
interface Turn {
  after?: string
  until?: string
  edit?: (context: TurnContext) => Promise<void>
  calls?: (context: TurnContext) => Array<{ id: string; name: string; args: unknown }>
  text?: string
}

const DATA: Array<Array<string | number>> = [
  ['name', 'score'],
  ['beta', 2],
  ['alpha', 1],
  ['delta', 4],
  ['gamma', 3],
]

function startStub(
  requests: ChatRequest[],
  turns: Turn[],
  page: () => Page,
): Promise<Server> {
  let sheetId = ''
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
    // The id is assigned by the app on import, so the stub learns it the way a
    // model does — from the text it is handed. The workbook overview at the end
    // of the user's message carries it (`Active sheet: Sheet1 (id=sheet-1, …)`),
    // so it is available on the very first turn; tool outputs repeat it.
    for (const m of body.messages) {
      const text = typeof m.content === 'string' ? m.content : JSON.stringify(m.content ?? '')
      const found = /\(id=([^,\s)]+)/.exec(text)
      if (found?.[1]) sheetId = found[1]
    }
    const context: TurnContext = { page: page(), sheetId }
    const turn =
      turns.find(
        (t) =>
          (t.after === undefined || answered(t.after)) && (t.until === undefined || !answered(t.until)),
      ) ?? {}
    await turn.edit?.(context)
    const message: ChatMessage = turn.text
      ? { role: 'assistant', content: turn.text }
      : {
          role: 'assistant',
          content: null,
          tool_calls: (turn.calls?.(context) ?? []).map((c) => ({
            id: c.id,
            type: 'function' as const,
            function: { name: c.name, arguments: JSON.stringify(c.args) },
          })),
        }
    res.writeHead(200, { 'content-type': 'application/json' })
    res.end(JSON.stringify({ choices: [{ message, finish_reason: 'stop' }] }))
  })
  return new Promise((r) => server.listen(0, '127.0.0.1', () => r(server)))
}

/** The tool output the app reported back for a given tool call id. */
function toolResult(requests: ChatRequest[], id: string): string {
  for (const request of requests) {
    for (const m of request.messages) {
      if (m.role === 'tool' && typeof m.content === 'string' && m.tool_call_id === id) return m.content
    }
  }
  return ''
}

async function run(
  testInfo: string,
  ask: string,
  turns: Turn[],
): Promise<{ toolResultOf: (id: string) => string; close: () => Promise<void> }> {
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-drift-'))
  const workbook = join(scratch, 'book.xlsx')
  await writeFile(workbook, await buildGridWorkbook(DATA))

  const requests: ChatRequest[] = []
  // The stub listens before the app exists (ai-settings.json is read at
  // launch), so it reaches the page through this accessor — no request can
  // arrive before the app has a window to send it from.
  let page: Page | undefined
  const server = await startStub(requests, turns, () => {
    if (!page) throw new Error('model request before the app was ready')
    return page
  })
  const port = (server.address() as AddressInfo).port

  // ai-settings.json is read at launch, so the stub has to be listening first.
  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-e2e-'))
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = { apiKey: 'stub', model: 'stub-model', baseUrl: `http://127.0.0.1:${port}/v1` }
  await writeFile(join(userDataDir, 'ai-settings.json'), JSON.stringify(resolveCustomEndpoints(settings)))

  const launched = await launchShell({
    userDataDir,
    onboardingSeen: true,
    openFile: workbook,
    videoDir: testInfo,
    // Exposes window.__genofficeDebug.univerAPI, which lands the user's edit.
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
    await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    await sheets.waitForTimeout(2_000)
    page = sheets

    const composer = sheets.getByPlaceholder(/Ask AI to analyze/)
    await composer.click()
    await composer.fill(ask)
    await sheets.getByRole('button', { name: 'Send', exact: true }).click()
    await expect(sheets.getByText('Done.').first()).toBeVisible({ timeout: 60_000 })
    await sheets.screenshot({ path: `/tmp/drift-${testInfo}.png` })
    return {
      toolResultOf: (id: string) => toolResult(requests, id),
      close: async () => {
        await launched.app.close().catch(() => {})
        server.close()
      },
    }
  } catch (error) {
    await launched.app.close().catch(() => {})
    server.close()
    throw error
  }
}

/**
 * A sort through the same command the Data ▸ Sort menu runs (the menu is a thin
 * wrapper over it), sorting rows 2–5 by column B ascending.
 */
async function sortRowsByScore(page: Page): Promise<void> {
  await page.evaluate(async () => {
    const debug = (
      window as unknown as {
        __genofficeDebug: {
          univerAPI: {
            getActiveWorkbook(): {
              getId(): string
              getActiveSheet(): { getSheetId(): string }
            }
            executeCommand(id: string, params: unknown): Promise<unknown>
          }
        }
      }
    ).__genofficeDebug
    const workbook = debug.univerAPI.getActiveWorkbook()
    await debug.univerAPI.executeCommand('sheet.command.sort-range', {
      unitId: workbook.getId(),
      subUnitId: workbook.getActiveSheet().getSheetId(),
      range: { startRow: 1, startColumn: 0, endRow: 4, endColumn: 1 },
      orderRules: [{ colIndex: 1, type: 'asc' }],
      hasTitle: true,
    })
  })
  await page.waitForTimeout(800)
}

/** ⌘Z with the grid focused, as a person would after seeing an AI change land. */
async function pressUndo(page: Page): Promise<void> {
  // Focus the sheet first: with the composer focused the shortcut would not be
  // the grid's to handle. D10 is outside everything either scenario writes.
  const spot = await gridCellCenter(page, 9, 3)
  await page.mouse.click(spot.x, spot.y)
  await page.waitForTimeout(300)
  await page.keyboard.press('ControlOrMeta+z')
  await page.waitForTimeout(800)
}

/**
 * A refusal has to be attributable. The gate's message and a schema rejection
 * both come back as errors, and a stub that sent a malformed plan would "pass"
 * a test about the gate while proving nothing — so every case asserts the plan
 * itself was accepted.
 */
function expectWellFormedPlan(output: string): void {
  expect(output).not.toContain('Invalid input')
  expect(output).not.toContain('unknown field')
}

test.describe('sheets AI: freshness after edits the model did not make', () => {
  /** The same `fill_range`, with and without a user sort in between. */
  async function fillCase(sortFirst: boolean): Promise<string> {
    const { toolResultOf, close } = await run(
      sortFirst ? 'drift-sort-fill' : 'drift-fill-control',
      'Extend the table by filling the last two rows from the row above.',
      [
        { until: 'read_1', calls: () => [{ id: 'read_1', name: 'read_range', args: { range: 'A1:B5' } }] },
        {
          after: 'read_1',
          until: 'fill_1',
          // The window: the model has read the block, the fill has not run.
          edit: sortFirst ? ({ page }) => sortRowsByScore(page) : undefined,
          calls: ({ sheetId }) => [
            {
              id: 'fill_1',
              name: 'propose_operations',
              args: {
                operations: [{ op: 'fill_range', sheetId, source: 'A2:B2', target: 'A4:B5' }],
                summary: 'extend the table',
              },
            },
          ],
        },
        { after: 'fill_1', text: 'Done.' },
      ],
    )
    try {
      return toolResultOf('fill_1')
    } finally {
      await close()
    }
  }

  test('a fill_range nobody moved still lands (control for the sort case)', async () => {
    const result = await fillCase(false)
    expectWellFormedPlan(result)
    expect(result).toContain('Auto-applied')
  })

  test('a manual sort blocks a fill_range over the sorted block', async () => {
    const result = await fillCase(true)
    expectWellFormedPlan(result)
    // A sort keeps every address but moves every record, so a bulk op over
    // the block is refused until the model re-reads.
    expect(result).toContain('structure changed')
    expect(result).toContain('get_workbook_context')
    expect(result).not.toContain('Auto-applied')
  })

  test('a write after the user ⌘Z-d the previous one stands, with a notice', async () => {
    const { toolResultOf, close } = await run(
      'drift-undo',
      'Label the total row at A6.',
      [
        { until: 'read_1', calls: () => [{ id: 'read_1', name: 'read_range', args: { range: 'A1:B5' } }] },
        {
          after: 'read_1',
          until: 'write_1',
          calls: ({ sheetId }) => [
            {
              id: 'write_1',
              name: 'propose_operations',
              args: {
                operations: [{ op: 'set_cell', sheetId, address: 'A6', value: 'total' }],
                summary: 'label A6',
              },
            },
          ],
        },
        {
          after: 'write_1',
          until: 'write_2',
          // The user takes the AI's label back. Addresses still valid, so the
          // next write must go through — but the model's picture of the sheet is
          // now wrong and has to be told.
          edit: ({ page }) => pressUndo(page),
          calls: ({ sheetId }) => [
            {
              id: 'write_2',
              name: 'propose_operations',
              args: {
                operations: [{ op: 'set_cell', sheetId, address: 'B6', value: 10 }],
                summary: 'put the total in B6',
              },
            },
          ],
        },
        { after: 'write_2', text: 'Done.' },
      ],
    )
    try {
      const first = toolResultOf('write_1')
      expectWellFormedPlan(first)
      expect(first).toContain('Auto-applied')
      const second = toolResultOf('write_2')
      // Content drift is reported, never enforced...
      expectWellFormedPlan(second)
      expect(second).toContain('Auto-applied')
      // ...and the report is what makes the model re-read instead of continuing
      // from numbers it no longer has.
      expect(second).toContain('contents changed since you last read')
    } finally {
      await close()
    }
  })
})
