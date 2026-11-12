import { expect, test } from '@playwright/test'
import { createServer, type Server } from 'node:http'
import { mkdir, writeFile } from 'node:fs/promises'
import { existsSync, mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { AddressInfo } from 'node:net'
import type { Page } from '@playwright/test'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { buildGridWorkbook, gridCellCenter, launchShell, waitForPageWithUrl } from './helpers'

/**
 * The fork's core fix — "AI notices what the user changed" — exercised in the
 * **packaged** Linux build, with the user's edits **stacked**.
 *
 * The single-edit cases already exist (`sheets-ai-stale-write.spec.ts` inserts a
 * row; `sheets-ai-drift.spec.ts` sorts and ⌘Zs). What they do not cover is the
 * combination the user asked for by name: several edits landing in the same
 * read→write window, including a structural edit *and* a content edit at once.
 * That combination is where a naive gate goes wrong in both directions — it can
 * file the whole thing as "just content" and silently write at coordinates the
 * user already moved, or it can over-block and refuse an edit that was fine.
 *
 * The gate (`apps/sheets/src/renderer/op-executor.ts:429`) checks
 * `workbookDrift().structural` **first**: any structural edit anywhere is a hard
 * refusal, regardless of content drift. Content drift alone is a soft notice and
 * the write still lands. D1–D6 pin both halves of that rule; D6 is the reverse
 * tooth — if the gate ever became "block on any drift", D6 goes red.
 *
 * Why packaged: this also proves the fork's renderer bundle (verified byte-present
 * in `linux-packaged-artifact.spec.ts` A3) actually *behaves* under the installed
 * layout, not just that the file is there.
 *
 * A scripted OpenAI-compatible stub makes the model's turns deterministic; the
 * user's edit lands between turns through the debug facade, which is how a real
 * ribbon/⌘Z edit reaches the same Univer commands. Evidence is read from the tool
 * output the app sends back to the model — the only channel the model itself sees.
 *
 * Gated so the shared suite is untouched on mac/win/CI:
 *   GENOFFICE_LINUX_SMOKE=1 DISPLAY=:77 npx playwright test \
 *     --config e2e/playwright.config.ts e2e/linux-freshness-packaged.spec.ts
 */

const ENABLED = process.platform === 'linux' && process.env.GENOFFICE_LINUX_SMOKE === '1'
const SHELL_DIR = resolve(__dirname, '../apps/shell')
const RELEASE = process.env.GENOFFICE_LINUX_DIST ?? join(SHELL_DIR, 'release')
const EXE = join(RELEASE, 'linux-unpacked/genoffice')

const freshHome = () => mkdtempSync(join(tmpdir(), 'genoffice-fresh-home-'))

const DATA: Array<Array<string | number>> = [
  ['name', 'score'],
  ['beta', 2],
  ['alpha', 1],
  ['delta', 4],
  ['gamma', 3],
]

interface ChatMessage {
  role: string
  content: unknown
  tool_call_id?: string
  tool_calls?: Array<{ id: string; type: 'function'; function: { name: string; arguments: string } }>
}
interface ChatRequest {
  messages: ChatMessage[]
}
interface TurnContext {
  page: Page
  sheetId: string
}
interface Turn {
  after?: string
  until?: string
  edit?: (context: TurnContext) => Promise<void>
  calls?: (context: TurnContext) => Array<{ id: string; name: string; args: unknown }>
  text?: string
}

function startStub(requests: ChatRequest[], turns: Turn[], page: () => Page): Promise<Server> {
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
    // The stub learns the import-assigned sheet id the way a model does — out of
    // the workbook overview text (`Active sheet: Sheet1 (id=sheet-1, …)`).
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

function toolResult(requests: ChatRequest[], id: string): string {
  for (const request of requests) {
    for (const m of request.messages) {
      if (m.role === 'tool' && typeof m.content === 'string' && m.tool_call_id === id) return m.content
    }
  }
  return ''
}

/**
 * Boots the packaged build against the stub. The AI settings go into the fake
 * home's profile (`$HOME/.config/GenOffice/ai-settings.json`) before launch —
 * `GENOFFICE_USER_DATA` is ignored once `app.isPackaged`, so HOME redirection is
 * the only isolation, and it is also what makes the packaged app read *our* stub.
 */
async function runPackaged(
  label: string,
  ask: string,
  turns: Turn[],
): Promise<{ toolResultOf: (id: string) => string; sheets: Page; close: () => Promise<void> }> {
  const homeDir = freshHome()
  const workbook = join(homeDir, 'book.xlsx')
  await writeFile(workbook, await buildGridWorkbook(DATA))

  const requests: ChatRequest[] = []
  let page: Page | undefined
  const server = await startStub(requests, turns, () => {
    if (!page) throw new Error('model request before the app was ready')
    return page
  })
  const port = (server.address() as AddressInfo).port

  const profileDir = join(homeDir, '.config', 'GenOffice')
  await mkdir(profileDir, { recursive: true })
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = { apiKey: 'stub', model: 'stub-model', baseUrl: `http://127.0.0.1:${port}/v1` }
  await writeFile(join(profileDir, 'ai-settings.json'), JSON.stringify(resolveCustomEndpoints(settings)))

  const launched = await launchShell({
    packaged: { executable: EXE, homeDir },
    onboardingSeen: true,
    openFile: workbook,
    settings: { starPrompt: { resolved: true } },
    videoDir: 'linux-freshness',
    // no_proxy so the stub on 127.0.0.1 is never routed through a host proxy.
    env: {
      GENOFFICE_DEBUG_HOOKS: '1',
      no_proxy: '127.0.0.1,localhost',
      NO_PROXY: '127.0.0.1,localhost',
    },
  })
  try {
    const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
    // Gate on real content, not the Sheet1 tab label (see linux-packaged-smoke B5).
    await expect
      .poll(
        () =>
          sheets
            .evaluate(() => {
              const debug = (
                window as unknown as {
                  __genofficeDebug: { univerAPI: { getActiveWorkbook(): any } }
                }
              ).__genofficeDebug
              return debug.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getRange(0, 0, 1, 1).getValue() ?? null
            })
            .catch(() => null),
        { timeout: 40_000, intervals: [500] },
      )
      .toBe('name')
    page = sheets

    const composer = sheets.getByPlaceholder(/Ask AI to analyze/)
    await composer.click()
    await composer.fill(ask)
    await sheets.getByRole('button', { name: 'Send', exact: true }).click()
    await expect(sheets.getByText('Done.').first()).toBeVisible({ timeout: 90_000 })
    await sheets.screenshot({ path: join(process.cwd(), `e2e/artifacts/screenshots/linux-D-${label}.png`) })
    return {
      toolResultOf: (id: string) => toolResult(requests, id),
      sheets,
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

// ── user edits, run through the debug facade (same Univer commands the ribbon
//    and ⌘Z drive) ────────────────────────────────────────────────────────────

/**
 * Run one Univer facade call on the active sheet, exactly as the ribbon does.
 * `call` is a method chain evaluated against the FWorksheet — e.g.
 * `insertRowsBefore(1, 1)` or `getRange('D2:E3').merge()`. These are the same
 * facade methods the app's own op-executor uses (`op-executor.ts:612-638`), so
 * a facade edit fires the same `sheet.mutation.*` the freshness gate watches.
 */
async function facadeEdit(page: Page, call: string): Promise<void> {
  await page.evaluate((chain) => {
    const debug = (
      window as unknown as { __genofficeDebug: { univerAPI: { getActiveWorkbook(): any } } }
    ).__genofficeDebug
    const ws = debug.univerAPI.getActiveWorkbook().getActiveSheet()
    // eslint-disable-next-line no-new-func
    return new Function('ws', `return ws.${chain}`)(ws)
  }, call)
  await page.waitForTimeout(700)
}

const insertRow = (page: Page, at = 1) => facadeEdit(page, `insertRowsBefore(${at}, 1)`)
const insertCol = (page: Page, at = 1) => facadeEdit(page, `insertColumnsBefore(${at}, 1)`)
const deleteRow = (page: Page, at = 2) => facadeEdit(page, `deleteRow(${at})`)
const mergeCells = (page: Page, range: string) => facadeEdit(page, `getRange('${range}').merge()`)
const addFilter = (page: Page, range: string) => facadeEdit(page, `getRange('${range}').createFilter()`)
const setValue = (page: Page, r: number, c: number, v: unknown) =>
  facadeEdit(page, `getRange(${r}, ${c}, 1, 1).setValue(${JSON.stringify(v)})`)

/** A sort through the same command the Data ▸ Sort menu runs (drift-spec precedent). */
async function sortRowsByScore(page: Page): Promise<void> {
  await page.evaluate(async () => {
    const debug = (
      window as unknown as {
        __genofficeDebug: {
          univerAPI: {
            getActiveWorkbook(): { getId(): string; getActiveSheet(): { getSheetId(): string } }
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

/** ⌘Z with the grid focused (D10 is outside every write in these scenarios). */
async function pressUndo(page: Page): Promise<void> {
  const spot = await gridCellCenter(page, 9, 3)
  await page.mouse.click(spot.x, spot.y)
  await page.waitForTimeout(300)
  await page.keyboard.press('ControlOrMeta+z')
  await page.waitForTimeout(800)
}

/** A refusal must be attributable: not a schema rejection masquerading as a gate. */
function expectWellFormedPlan(output: string): void {
  expect(output).not.toContain('Invalid input')
  expect(output).not.toContain('unknown field')
}

// A read turn, then a write turn to the address the read established. The user
// edit lands in the `edit` of the write turn — between the read and the write.
function readThenWrite(edit: ((c: TurnContext) => Promise<void>) | undefined, writeId = 'write_1') {
  return [
    { until: 'read_1', calls: () => [{ id: 'read_1', name: 'read_range', args: { range: 'A1:B5' } }] },
    {
      after: 'read_1',
      until: writeId,
      edit,
      calls: ({ sheetId }: TurnContext) => [
        {
          id: writeId,
          name: 'propose_operations',
          args: {
            operations: [{ op: 'set_cell', sheetId, address: 'A1', value: 'WRITTEN' }],
            summary: 'label A1',
          },
        },
      ],
    },
    { after: writeId, text: 'Done.' },
  ] as Turn[]
}

const STRUCTURAL = 'structure changed'
const REASK = 'get_workbook_context'

test.describe('linux packaged: AI freshness under stacked user edits', () => {
  test.skip(!ENABLED, 'set GENOFFICE_LINUX_SMOKE=1 (linux only)')
  test.skip(!existsSync(EXE), `no packaged build at ${EXE} — run npm run dist:linux`)

  test('D1 insert row + change a value → hard block (structure wins over content)', async () => {
    test.setTimeout(180_000)
    const { toolResultOf, close } = await runPackaged(
      'D1',
      'Put the label WRITTEN in A1.',
      readThenWrite(async ({ page }) => {
        await insertRow(page, 1) // structural
        await setValue(page, 3, 1, 99) // content, in the same window
      }),
    )
    try {
      const result = toolResultOf('write_1')
      expectWellFormedPlan(result)
      // The whole point of stacking: a content edit must not downgrade the
      // structural refusal to a soft notice.
      expect(result).toContain(STRUCTURAL)
      expect(result).toContain(REASK)
      expect(result).not.toContain('Auto-applied')
    } finally {
      await close()
    }
  })

  test('D2 insert column + delete row (two structural edits) → hard block', async () => {
    test.setTimeout(180_000)
    const { toolResultOf, close } = await runPackaged(
      'D2',
      'Put the label WRITTEN in A1.',
      readThenWrite(async ({ page }) => {
        await insertCol(page, 1)
        await deleteRow(page, 2)
      }),
    )
    try {
      const result = toolResultOf('write_1')
      expectWellFormedPlan(result)
      expect(result).toContain(STRUCTURAL)
      expect(result).not.toContain('Auto-applied')
    } finally {
      await close()
    }
  })

  test('D3 manual sort + change a block value, then fill_range → hard block', async () => {
    test.setTimeout(180_000)
    const { toolResultOf, close } = await runPackaged(
      'D3',
      'Extend the table by filling the last two rows from the row above.',
      [
        { until: 'read_1', calls: () => [{ id: 'read_1', name: 'read_range', args: { range: 'A1:B5' } }] },
        {
          after: 'read_1',
          until: 'fill_1',
          edit: async ({ page }) => {
            await sortRowsByScore(page) // structural: moves every record
            await setValue(page, 2, 1, 77) // content, inside the block
          },
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
      const result = toolResultOf('fill_1')
      expectWellFormedPlan(result)
      // A sort keeps every address but moves every record; a bulk op whose
      // meaning depends on block contents is refused until a re-read.
      expect(result).toContain(STRUCTURAL)
      expect(result).not.toContain('Auto-applied')
    } finally {
      await close()
    }
  })

  test('D4 AI write → user ⌘Z → user insert row → second write → hard block', async () => {
    test.setTimeout(200_000)
    const { toolResultOf, close } = await runPackaged(
      'D4',
      'Label the total row, then put the total in B6.',
      [
        { until: 'read_1', calls: () => [{ id: 'read_1', name: 'read_range', args: { range: 'A1:B5' } }] },
        {
          after: 'read_1',
          until: 'write_1',
          calls: ({ sheetId }) => [
            {
              id: 'write_1',
              name: 'propose_operations',
              args: { operations: [{ op: 'set_cell', sheetId, address: 'A6', value: 'total' }], summary: 'label A6' },
            },
          ],
        },
        {
          after: 'write_1',
          until: 'write_2',
          // Content drift (the ⌘Z) AND structural drift (the row insert) coexist.
          // Content alone would be a soft notice; the insert must escalate to a
          // hard refusal.
          edit: async ({ page }) => {
            await pressUndo(page)
            await insertRow(page, 1)
          },
          calls: ({ sheetId }) => [
            {
              id: 'write_2',
              name: 'propose_operations',
              args: { operations: [{ op: 'set_cell', sheetId, address: 'B6', value: 10 }], summary: 'total in B6' },
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
      expectWellFormedPlan(second)
      expect(second).toContain(STRUCTURAL)
      expect(second).not.toContain('Auto-applied')
    } finally {
      await close()
    }
  })

  test('D5 merge cells (+ add filter) → hard block (merge is structural)', async () => {
    test.setTimeout(180_000)
    const { toolResultOf, close } = await runPackaged(
      'D5',
      'Put the label WRITTEN in A1.',
      readThenWrite(async ({ page }) => {
        await mergeCells(page, 'D2:E3') // structural, outside the read range
        await addFilter(page, 'A1:B5') // content (filter)
      }),
    )
    try {
      const result = toolResultOf('write_1')
      expectWellFormedPlan(result)
      expect(result).toContain(STRUCTURAL)
      expect(result).not.toContain('Auto-applied')
    } finally {
      await close()
    }
  })

  test('D6 reverse tooth: a content-only edit outside the read range still lands', async () => {
    test.setTimeout(180_000)
    const { toolResultOf, close } = await runPackaged(
      'D6',
      'Put the label WRITTEN in A1.',
      readThenWrite(({ page }) => setValue(page, 9, 3, 'note outside the table')),
    )
    try {
      const result = toolResultOf('write_1')
      expectWellFormedPlan(result)
      // No structural edit → the write must go through. If the gate ever became
      // "refuse on any drift", this is the case that turns red.
      expect(result).toContain('Auto-applied')
      expect(result).not.toContain(STRUCTURAL)
      // Content drift is still reported so the model re-reads before quoting old numbers.
      expect(result).toContain('contents changed since you last read')
    } finally {
      await close()
    }
  })
})
