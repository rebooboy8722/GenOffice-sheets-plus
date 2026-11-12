import { expect, test } from '@playwright/test'
import { mkdir, mkdtemp, readdir, readFile, writeFile } from 'node:fs/promises'
import { execSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Page } from '@playwright/test'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { ARTIFACTS_DIR, buildGridWorkbook, gridCellCenter, launchShell, waitForPageWithUrl } from './helpers'

/**
 * A real-endpoint regression net for two things a stub cannot reach, both taken
 * from what the user actually hit in the shipped build (2026-10-24):
 *
 *  1. **A bulk write lands.** The reported failure was a rejected batch — the
 *     model sent `set_range` with invented fields (`startCell`/`data`; the real
 *     ones are `start`/`values`), and a batch is all-or-nothing, so nothing was
 *     written and a whole round trip was burned.
 *  2. **A manual edit reaches the model.** The user's own summary of the whole
 *     fork: "这次最大的意义应该是手动修改让模型知道". The user inserts a row by
 *     hand between turns, then asks for a write *relative to a value* ("next to
 *     the row that says gamma"). Every address the model holds is now off by one.
 *     If nothing tells it, the word lands next to `beta` — a silent miswrite,
 *     which is the defect this fork exists to fix. The assertion is semantic
 *     (same row as `gamma`) so a model that re-reads on its own and one that is
 *     refused by the freshness gate both pass, and only the silent miswrite
 *     fails.
 *
 * **What this spec does NOT prove — read before citing it as evidence.** It is
 * not the teeth for the `set_range` description fix (`7f5958e3`). Reverting that
 * fix and re-running leaves this green: whether the model sends `set_range` at
 * all, and whether it guesses the field names, is stochastic, and across four
 * runs without the fix it never once produced `startCell`/`data`. The
 * deterministic check for that fix is
 * `apps/sheets/tests/tool-description-set-range.test.ts`. What this spec earns
 * its keep on is the end-to-end path — real xlsx through the sidecar, real
 * model, a real user edit between turns, and the result verified in the live
 * grid and in the saved file.
 *
 * Evidence comes from the AI trace (one JSONL line per provider turn: tool calls
 * under `response.toolCalls[].input`, tool results, tokens, wall time) rather
 * than from model prose, which is model-chosen. The summary — including the ops
 * the model actually sent, so a green run can be read for *what it proved* —
 * goes to `e2e/artifacts/real-write-<label>.json` for comparison across builds.
 *
 * Needs the local model endpoint; skips when it is unreachable.
 */

const BASE_URL = process.env.GENOFFICE_E2E_AI_URL ?? 'http://127.0.0.1:8080/v1'
const MODEL = process.env.GENOFFICE_E2E_AI_MODEL ?? 'Qwen3.6-27B-UD-Q5_K_XL.gguf'
const LABEL =
  process.env.GENOFFICE_E2E_TURN_LABEL ?? new Date().toISOString().replace(/[:.]/g, '')
/**
 * A literal 4×4 block of text — chosen so `set_range` is the only sensible op.
 *
 * Two earlier drafts were toothless and both were caught the same way: by reading
 * **what the model actually sent** out of the trace, not the test verdict.
 *   - "write low/medium/high into C2:C4" → 3× `set_cell`. Correct answer, zero
 *     `set_range`, so the `set_range` description fix was never exercised.
 *   - "a times-table in B7:D10" → 1× `set_range` for column B plus 8×
 *     `set_formula` (`=B7*2`). Again a *better* answer than the one I wanted —
 *     formulas beat hardcoded numbers — and with the fix reverted the model
 *     answered the same task with 12 `set_cell`/`set_formula` ops and the test
 *     still passed. A numeric relationship invites formulas; a block of prose
 *     does not, and 16 `set_cell` ops is not a shape models reach for when a
 *     2D `values` array exists.
 *
 * The assertion that gives this teeth is `rejected1 === 0`: a batch is
 * all-or-nothing, so invented field names (`startCell`/`data`) fail the whole
 * batch and burn a round trip — the exact failure the user reported.
 */
const WRITE_TASK =
  'Write this 4x4 block of text into B7:E10 exactly as written here, one item per cell, ' +
  'row by row: north, south, east, west / kilo, lima, oscar, papa ' +
  '/ red, green, blue, amber / swift, brisk, calm, keen. ' +
  'These are plain text labels, not formulas. Do not change any other cell.'
/**
 * The block in the order the task states it: rows B7:E10.
 * Deliberately shares no word with the fixture (name/qty/alpha/beta/gamma) — the
 * drift turn locates its row with `find_cells("gamma")`, and a second `gamma`
 * written here would make that ambiguous and sink scenario ② for the wrong reason.
 */
const BLOCK = [
  ['north', 'south', 'east', 'west'],
  ['kilo', 'lima', 'oscar', 'papa'],
  ['red', 'green', 'blue', 'amber'],
  ['swift', 'brisk', 'calm', 'keen'],
]
const MARKER = 'ZZZMARK'
// Say "replace whatever is there", not "the empty cell to the right": the cell
// right of gamma holds a value in the fixture, so calling it "empty" made the
// task unsatisfiable-as-written. In one run the model spotted exactly that and
// spent 2703 reasoning chars deciding to ask instead of acting — a correct read
// of a bad prompt, scored by this spec as a failure. An oracle must not be
// answerable with "your question is broken".
const DRIFT_TASK =
  `Write the word ${MARKER} into the cell immediately to the right of the cell that contains ` +
  'the text gamma, replacing whatever that cell currently holds. Change nothing else.'

interface DebugWindow {
  __genofficeDebug: {
    univerAPI: {
      getActiveWorkbook(): {
        getActiveSheet(): {
          getRange(r: number, c: number, h?: number, w?: number): { getValue(): unknown }
        } | null
      } | null
    }
  }
}

/** A turn's trace line, as written by packages/ai-provider/src/trace.ts. */
interface TraceLine {
  durationMs: number
  messages: Array<{
    role: string
    text?: string
    results?: Array<{ toolName?: string; output?: unknown; isError?: boolean }>
  }>
  response: {
    text?: string
    /** trace.ts stores the parsed call under `input`, not `arguments` — reading
     * the wrong key silently yields `[]` and the run looks clean for nothing. */
    toolCalls: Array<{ name: string; input?: unknown }>
    stopReason?: string
    error?: string
    usage?: { inputTokens?: number; outputTokens?: number }
  }
}

/** One cell as text — numbers and strings compared as text, empty as ''. */
function cellText(page: Page, row: number, col: number): Promise<string> {
  return page.evaluate(
    ({ row, col }) => {
      const win = window as unknown as DebugWindow
      const v = win.__genofficeDebug.univerAPI
        .getActiveWorkbook()
        ?.getActiveSheet()
        ?.getRange(row, col, 1, 1)
        .getValue()
      return v === null || v === undefined ? '' : String(v)
    },
    { row, col },
  )
}

/** Reads a strip of the grid: enough to locate `gamma` and the marker anywhere. */
function readGrid(page: Page, rows: number, cols: number): Promise<(string | null)[][]> {
  return page.evaluate(
    ({ rows, cols }) => {
      const win = window as unknown as DebugWindow
      const sheet = win.__genofficeDebug.univerAPI.getActiveWorkbook()?.getActiveSheet()
      const out: (string | null)[][] = []
      for (let r = 0; r < rows; r++) {
        const line: (string | null)[] = []
        for (let c = 0; c < cols; c++) {
          const v = sheet?.getRange(r, c, 1, 1).getValue()
          line.push(v === null || v === undefined ? null : String(v))
        }
        out.push(line)
      }
      return out
    },
    { rows, cols },
  )
}

async function readTrace(dir: string): Promise<TraceLine[]> {
  const files = (await readdir(dir).catch(() => [] as string[])).filter((f) => f.endsWith('.jsonl'))
  const lines: TraceLine[] = []
  for (const file of files) {
    const raw = await readFile(join(dir, file), 'utf8').catch(() => '')
    for (const line of raw.split('\n')) {
      if (!line.trim()) continue
      try {
        lines.push(JSON.parse(line) as TraceLine)
      } catch {
        /* a partially written line; the poll retries */
      }
    }
  }
  return lines
}

/**
 * A run is over when the last turn asked for no tools (the loop's terminal turn)
 * and the trace has been quiet. Model-agnostic, unlike waiting on prose; the
 * Stop button is a second signal because a model can also end on a tool error.
 */
async function waitForRunEnd(
  dir: string,
  page: Page,
  turnOffset: number,
  timeoutMs = 300_000,
): Promise<TraceLine[]> {
  const deadline = Date.now() + timeoutMs
  let lastChange = Date.now()
  let lastCount = -1
  for (;;) {
    const lines = await readTrace(dir)
    if (lines.length !== lastCount) {
      lastCount = lines.length
      lastChange = Date.now()
    }
    const idle =
      (await page.getByRole('button', { name: 'Stop', exact: true }).count()) === 0
    const last = lines.at(-1)
    if (
      lines.length > turnOffset &&
      last &&
      last.response.toolCalls.length === 0 &&
      idle &&
      Date.now() - lastChange > 3_000
    ) {
      return lines
    }
    if (Date.now() > deadline) return lines
    await new Promise((r) => setTimeout(r, 1_000))
  }
}

/** Parts are deflate-compressed, so search the unzipped sheet, not the zip. */
const savedSheet = (path: string): string => {
  try {
    return execSync(`unzip -p "${path}" xl/worksheets/sheet1.xml`).toString()
  } catch {
    return ''
  }
}

const flatten = (lines: TraceLine[]): string[] =>
  lines.flatMap((l) => (l.messages ?? []).flatMap((m) => (m.results ?? []).map((r) => String(r.output ?? ''))))

test('a real bulk write lands, and a manual row insert is not written past', async () => {
  const reachable = await fetch(`${BASE_URL.replace(/\/v1$/, '')}/v1/models`, {
    signal: AbortSignal.timeout(6_000),
  })
    .then((r) => r.ok)
    .catch(() => false)
  test.skip(!reachable, `model endpoint unreachable at ${BASE_URL}`)
  test.setTimeout(600_000)

  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-real-write-'))
  const workbook = join(scratch, 'book.xlsx')
  // A real file on disk: the lazy sidecar path the user is on, not a demo workbook.
  await writeFile(
    workbook,
    await buildGridWorkbook([
      ['name', 'qty'],
      ['alpha', 1],
      ['beta', 2],
      ['gamma', 3],
    ]),
  )
  const traceDir = join(scratch, 'trace')
  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-e2e-'))

  const settings = defaultAiSettings()
  settings.provider = 'custom'
  // GENOFFICE_E2E_EFFORT sweeps the endpoint's reasoning effort through the only
  // channel that reaches the wire today (profile.bodyExtras, spread verbatim).
  // 'off' disables thinking; unset leaves the endpoint default (xhigh for this
  // model) untouched so the run doubles as the control arm.
  const effort = process.env.GENOFFICE_E2E_EFFORT
  const chatTemplateKwargs =
    effort === 'off' ? { enable_thinking: false } : effort ? { reasoning_effort: effort } : undefined
  settings.providers.custom = {
    apiKey: 'local',
    model: MODEL,
    baseUrl: BASE_URL,
    ...(chatTemplateKwargs ? { profile: { bodyExtras: { chat_template_kwargs: chatTemplateKwargs } } } : {}),
  }
  await writeFile(join(userDataDir, 'ai-settings.json'), JSON.stringify(resolveCustomEndpoints(settings)))

  const summary: Record<string, unknown> = {
    label: LABEL,
    baseUrl: BASE_URL,
    model: MODEL,
    effort: process.env.GENOFFICE_E2E_EFFORT ?? '(endpoint default)',
    writeTask: WRITE_TASK,
    driftTask: DRIFT_TASK,
  }

  const launched = await launchShell({
    userDataDir,
    onboardingSeen: true,
    openFile: workbook,
    videoDir: 'real-write',
    env: { GENOFFICE_DEBUG_HOOKS: '1', GENOFFICE_AI_TRACE_DIR: traceDir },
  })
  try {
    const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
    await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    await sheets.waitForTimeout(2_000)

    const composer = sheets.getByPlaceholder(/Ask AI to analyze/)

    // ── 1. bulk write ────────────────────────────────────────────────────────
    await composer.click()
    await composer.fill(WRITE_TASK)
    await sheets.getByRole('button', { name: 'Send', exact: true }).click()
    const firstRun = await waitForRunEnd(traceDir, sheets, 0)
    summary.turns1 = firstRun.length
    summary.tools1 = firstRun.map((l) => l.response.toolCalls.map((c) => c.name))
    summary.rejected1 = flatten(firstRun).filter((o) => /Rejected — none of/i.test(o)).length
    summary.rejectionTexts1 = flatten(firstRun)
      .filter((o) => /Rejected — none of/i.test(o))
      .map((o) => o.slice(0, 500))

    // Every op the model actually proposed, so a green run can be read for
    // *what it proves*: a run answered entirely with set_cell exercised nothing
    // about the set_range description fix.
    summary.ops1 = firstRun
      .flatMap((l) => l.response.toolCalls)
      .flatMap((c) => {
        const input = c.input as { operations?: Array<{ op?: string }> } | undefined
        return (input?.operations ?? []).map((o) => o.op ?? '?')
      })
    summary.setRangeInput = firstRun
      .flatMap((l) => l.response.toolCalls)
      .map((c) => c.input)
      .filter((i) => JSON.stringify(i ?? '').includes('set_range'))

    // Every cell of the block, exactly. `toBe` rather than a substring match so a
    // wrong-but-containing value can't pass, and `rejected1 === 0` is the tooth:
    // a batch is all-or-nothing, so one invented field name fails all 16 cells.
    for (const [i, row] of BLOCK.entries()) {
      for (const [j, expected] of row.entries()) {
        await expect
          .poll(() => cellText(sheets, 6 + i, 1 + j), { timeout: 20_000 })
          .toBe(expected)
      }
    }
    expect(summary.rejected1).toBe(0)
    // Read the artifact, not just the verdict: a green run that answered entirely
    // with set_cell exercised nothing about the set_range description.
    summary.usedSetRange = Array.isArray(summary.ops1) && summary.ops1.includes('set_range')

    // ── 2. the user edits by hand, then talks again ──────────────────────────
    // A row inserted at the top is *structural*: every address the model holds
    // from turn 1 is now one row off, so `gamma` sits in row 5, not row 4.
    // Same affordance sheets-ai-stale-write.spec.ts uses: click the row-1 header
    // (column -1) so the ribbon shows "Insert row", then click it. A keyboard
    // shortcut is layout-dependent and has silently no-opped before.
    const header = await gridCellCenter(sheets, 0, -1)
    await sheets.mouse.click(header.x, header.y)
    await sheets.waitForTimeout(300)
    await sheets.getByRole('button', { name: 'Insert row', exact: true }).click()
    await sheets.waitForTimeout(1_000)
    // The insert must really have happened for the second turn to mean anything:
    // row 1 (index 0) is now empty and the header has moved down.
    expect(await cellText(sheets, 0, 0)).toBeFalsy()
    expect(await cellText(sheets, 1, 0)).toMatch(/name/i)
    // 12 rows: the text block occupies rows 7–10, so a strip of 8 would clip it
    // and a marker written below the block would go unnoticed.
    const beforeDrift = await readGrid(sheets, 12, 6)
    const gammaRowAtTurn2 = beforeDrift.findIndex((line) => line.some((v) => /gamma/i.test(v ?? '')))
    expect(gammaRowAtTurn2).toBeGreaterThan(0)

    await composer.click()
    await composer.fill(DRIFT_TASK)
    await sheets.getByRole('button', { name: 'Send', exact: true }).click()
    const secondRun = await waitForRunEnd(traceDir, sheets, firstRun.length)
    const newLines = secondRun.slice(firstRun.length)
    summary.turns2 = newLines.length
    summary.tools2 = newLines.map((l) => l.response.toolCalls.map((c) => c.name))
    // Did the freshness gate fire at all? Diagnostic, not an assertion: a model
    // that re-reads on its own never trips it, and that is also a pass.
    summary.gateFired = flatten(newLines).some((o) =>
      /workbook structure changed since you last read it/i.test(o),
    )
    summary.rejected2 = flatten(newLines).filter((o) => /Rejected — none of/i.test(o)).length

    const after = await readGrid(sheets, 12, 6)
    const markerRows = after
      .map((line, i) => ({ i, hit: line.some((v) => v?.includes(MARKER)) }))
      .filter((r) => r.hit)
      .map((r) => r.i)
    const gammaRowAfter = after.findIndex((line) => line.some((v) => /gamma/i.test(v ?? '')))
    summary.grid = { gammaRowAtTurn2, gammaRowAfter, markerRows }
    await sheets.screenshot({ path: join(ARTIFACTS_DIR, 'real-write', 'after-drift-turn.png') }).catch(() => {})

    // The whole point: the marker lands on gamma's *current* row, exactly once.
    // The user's insert moved gamma down one row since turn 1, so a model still
    // holding its pre-insert addresses writes one row too high — landing on
    // gammaRowAfter - 1 is precisely the silent miswrite this fork exists to
    // prevent, and this assertion is what catches it.
    expect(markerRows).toHaveLength(1)
    expect(markerRows[0]).toBe(gammaRowAfter)
    // Row alone was too weak: it let a marker anywhere in gamma's row pass, so
    // it never proved the model resolved the *address*, only the row. Pin the
    // column too — one cell to the right of gamma.
    const gammaColAfter = after[gammaRowAfter].findIndex((v) => /gamma/i.test(v ?? ''))
    const markerColAfter = after[markerRows[0]!].findIndex((v) => v?.includes(MARKER))
    summary.markerCell = { row: markerRows[0], col: markerColAfter }
    expect(markerColAfter).toBe(gammaColAfter + 1)

    // And it must survive a save at that row — the grid is not the deliverable.
    await launched.app.evaluate(({ webContents }) => {
      const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('://sheets/'))
      wc?.send('menu:action', 'save')
    })
    await expect
      .poll(() => savedSheet(workbook), { timeout: 30_000 })
      .toContain(MARKER)
  } finally {
    // The summary is written even when an assertion failed: a red run is exactly
    // when the evidence about what the model sent matters most.
    summary.traceLines = (await readTrace(traceDir).catch(() => [] as TraceLine[])).length
    await mkdir(join(ARTIFACTS_DIR, 'real-write'), { recursive: true })
    await writeFile(
      join(ARTIFACTS_DIR, `real-write-${LABEL}.json`),
      JSON.stringify(summary, null, 2),
    ).catch(() => {})
    console.log('REAL_WRITE>>>', JSON.stringify(summary, null, 1), '<<<END')
    await launched.app.close().catch(() => {})
  }
})
