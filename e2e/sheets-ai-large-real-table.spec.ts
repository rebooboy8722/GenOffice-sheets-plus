import { expect, test } from '@playwright/test'
import { copyFile, mkdir, mkdtemp, readdir, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Page } from '@playwright/test'
import {
  defaultAiSettings,
  resolveCustomEndpoints,
  type AiCustomProfile,
} from '@genoffice/ai-provider'
import { ARTIFACTS_DIR, launchShell, waitForPageWithUrl } from './helpers'

/**
 * Does the agent stay *correct* on a big, real, interdependent table — and what
 * does it cost in thinking?
 *
 * The user's ask (2026-10-24): "你找个大的真实表格，数据比较多，而且逻辑关联的测试测试，
 * 就能发现大概问题". Everything measured so far ran on a hand-built 4×5 fixture,
 * which cannot reproduce the behaviour being complained about ("AI 在雷霆大思考，
 * 各自对比校验") — that appears precisely when there is a lot of data to cross-check.
 *
 * The table is a real export the user owns: 1257 records, 35 columns, with money /
 * date / category columns that are only *partially* filled (费用来源 blank in 497
 * rows, 预估费 blank in 326). Real dirt matters — blank handling is where a
 * summarising model quietly goes wrong, and no synthetic fixture has its shape.
 *
 * It is copied into a scratch dir and never opened in place: the model writes into
 * whatever file it is given, and this one is somebody's business data.
 *
 * The four questions are cross-column aggregates (a two-condition count, a sum over
 * a column with blanks, a threshold count, a date-year count). They can only be
 * answered by actually reading the sheet — a model that extrapolates from the first
 * screenful is wrong — while the *output* is four cells. Heavy read, light write:
 * that isolates reasoning and retrieval from transport, which a 693-cell marking
 * task would not.
 *
 * Every expected value was computed from the file by an independent script, never
 * by the model, and is pinned next to a fingerprint of the table. **The fingerprint
 * is the point.** This rig has twice produced a meaningless result because the key
 * did not match the data: once a hand-computed key that was simply wrong (all four
 * arms scored 0/5, measuring the key rather than the model), once a task whose
 * wording contradicted the fixture (the model spent 2703 reasoning chars correctly
 * deciding to ask, and was scored a failure). If the file is not the one these keys
 * came from, this spec fails loudly instead of reporting a precision loss that
 * belongs to its own bookkeeping.
 *
 * Needs the local model endpoint and `GENOFFICE_E2E_XLSX` pointing at that export
 * (kept out of git — it is personal data); skips otherwise.
 */

const BASE_URL = process.env.GENOFFICE_E2E_AI_URL ?? 'http://127.0.0.1:8080/v1'
const MODEL = process.env.GENOFFICE_E2E_AI_MODEL ?? 'Qwen3.6-27B-UD-Q5_K_XL.gguf'
const SRC = process.env.GENOFFICE_E2E_XLSX ?? ''
const LABEL =
  process.env.GENOFFICE_E2E_TURN_LABEL ?? new Date().toISOString().replace(/[:.]/g, '')
/**
 * Declare the endpoint's context window (tokens) in the model profile. Unset,
 * the run keeps the historical shape: no window declared, so the agent's
 * compaction budget falls back to the shipped 256KB default and the
 * window-calibrated budget path (the one calibration under test acts in) is
 * never taken. Set to the endpoint's real max_model_len (210 vLLM: 262144) to
 * reproduce the production shape — a probed/declared window — that the
 * bytes-per-token calibration exists to fix.
 */
const CONTEXT_WINDOW = Number(process.env.GENOFFICE_E2E_CONTEXT_WINDOW ?? '')

/** Fingerprint of the exact file the keys below were derived from. */
const EXPECT_ROWS = 1258
const EXPECT_HEADER = ['序号', '格式类型', '创建日期', '发起人', '主题', '外委类型', '项目类型']
/**
 * Where the four answers go: column AH (0-based 33), rows 2–5.
 *
 * AH is *inside* the grid and empty in all 1257 data rows (measured from the file:
 * Z, AH and AI are blank below their headers), so writing there is a legal write
 * and the task's "change nothing else" stays satisfiable.
 *
 * The first drafts of this spec asked for AJ instead — the empty column one past
 * the last data column — which is what a person would ask for, and which this
 * sheet cannot honour: the imported grid is exactly the data extent (1258 × 35),
 * so AJ does not exist. Asked for AJ, the model burned 38–46 rejected batches and
 * then restructured the table to make the request legal (one run moved 1258 cells,
 * another inserted a column and restored the displaced header). That is the
 * Tier 0 finding, and it belongs in a report, not in the task text: an oracle
 * whose only passing path is "violate the constraint you just stated" measures
 * the oracle. Set `GENOFFICE_E2E_OUTSIDE=1` to reproduce the rejection storm on
 * purpose; the collateral assertions are skipped in that mode, since a structural
 * change is then the model's only way to answer at all.
 */
const OUTSIDE = process.env.GENOFFICE_E2E_OUTSIDE === '1'
/** 0-based column index of the answer column: AH inside the grid, AJ outside it. */
const OUT_COL = OUTSIDE ? 35 : 33
/** 0-based row of the first answer: row 2, so the header row stays untouched. */
const OUT_ROW = 1
const OUT_LABEL = OUTSIDE ? 'AJ' : 'AH'

/** Computed from the source file by an independent script. */
const KEYS = {
  /** rows where 外委类型 = 工程外委 AND 预估费 is blank */
  conditionalBlank: 125,
  /** sum of 预估费, blanks ignored */
  sum: 119151200,
  /** rows with 预估费 > 1000000 */
  over1m: 25,
  /** rows whose 创建日期 falls in 2026 */
  year2026: 258,
}

const TASK =
  '这个工作表有 1257 条数据（第 2 行到第 1258 行），共 35 列。请统计出四个数字，' +
  `分别写成纯数字放到 ${OUT_LABEL} 列的第 2 到第 5 行：` +
  `${OUT_LABEL}2 = F 列(外委类型)等于"工程外委" 且 N 列(预估费)为空的行数；` +
  `${OUT_LABEL}3 = N 列(预估费)所有数值的总和（空单元格按 0 算）；` +
  `${OUT_LABEL}4 = N 列(预估费)大于 1000000 的行数；` +
  `${OUT_LABEL}5 = C 列(创建日期)是 2026 年的行数。` +
  '四个数字必须基于全部 1257 行统计，不要只看屏幕前面几行就下结论。' +
  '不要改动其他任何单元格，也不要增删行列。'

interface DebugWindow {
  __genofficeDebug: {
    univerAPI: {
      getActiveWorkbook(): {
        getActiveSheet(): {
          getRange(r: number, c: number, h?: number, w?: number): { getValue(): unknown }
          getMaxRows(): number
          getMaxColumns(): number
        } | null
      } | null
    }
  }
}

interface TraceLine {
  durationMs: number
  messages: Array<{
    role: string
    text?: string
    results?: Array<{ toolName?: string; output?: unknown; isError?: boolean }>
  }>
  response: {
    text?: string
    reasoning?: string
    /** stored under `input`, not `arguments` — the wrong key yields [] silently. */
    toolCalls: Array<{ name: string; input?: unknown }>
    stopReason?: string
    error?: string
    usage?: { inputTokens?: number; outputTokens?: number }
  }
}

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

function safeParse(value: string): unknown {
  try {
    return JSON.parse(value)
  } catch {
    return undefined
  }
}

/**
 * The workbook's sheet names — proves whether a scratch working sheet leaked
 * into the saved workbook after the run. The working-sheet guide tells the model
 * to delete its scratch tab, so a clean run ends with only the original sheet(s).
 */
function sheetNames(page: Page): Promise<string[]> {
  return page
    .evaluate(() => {
      const win = window as unknown as DebugWindow
      const wb = win.__genofficeDebug.univerAPI?.getActiveWorkbook?.() as
        | {
            getSheets?(): Array<
              | { name?: string; sheetName?: string }
              | { getSheetName?(): string }
            >
            | undefined
          }
        | undefined
      const sheets = wb?.getSheets?.() ?? []
      return sheets.map((s) => {
        if (typeof s === 'object' && s !== null) {
          if (typeof (s as { getSheetName?: () => string }).getSheetName === 'function')
            return (s as { getSheetName: () => string }).getSheetName()
          return (s as { name?: string; sheetName?: string }).name ?? (s as { sheetName?: string }).sheetName ?? ''
        }
        return String(s)
      })
    })
    .catch(() => [])
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
        /* partially written; the poll retries */
      }
    }
  }
  return lines
}

/** See the other spec: a run ends on a tool-less turn plus a quiet trace. */
async function waitForRunEnd(
  dir: string,
  page: Page,
  turnOffset: number,
  timeoutMs = 900_000,
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
    const idle = (await page.getByRole('button', { name: 'Stop', exact: true }).count()) === 0
    const last = lines.at(-1)
    if (
      lines.length > turnOffset &&
      last &&
      last.response.toolCalls.length === 0 &&
      idle &&
      Date.now() - lastChange > 5_000
    ) {
      return lines
    }
    if (Date.now() > deadline) return lines
    await new Promise((r) => setTimeout(r, 1_000))
  }
}

/** Normalises "119151200", "119151200.0" and "119,151,200". */
const asNumber = (s: string): number | null => {
  const t = s.replace(/[,，\s]/g, '')
  if (!/^-?\d+(\.\d+)?$/.test(t)) return null
  return Number(t)
}

/** 0-based column index → label, for naming leftover cells in failures. */
const columnOf = (index: number): string => {
  let n = index
  let out = ''
  do {
    out = String.fromCharCode(65 + (n % 26)) + out
    n = Math.floor(n / 26) - 1
  } while (n >= 0)
  return out
}

test('four cross-column aggregates over a 1257-row real table come out exact', async () => {
  test.skip(!SRC, 'set GENOFFICE_E2E_XLSX to the real export (kept out of git)')
  const reachable = await fetch(`${BASE_URL.replace(/\/v1$/, '')}/v1/models`, {
    signal: AbortSignal.timeout(6_000),
  })
    .then((r) => r.ok)
    .catch(() => false)
  test.skip(!reachable, `model endpoint unreachable at ${BASE_URL}`)
  test.setTimeout(1_500_000)

  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-bigtable-'))
  const workbook = join(scratch, 'book.xlsx')
  await copyFile(SRC, workbook)
  const traceDir = join(scratch, 'trace')
  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-e2e-'))

  const settings = defaultAiSettings()
  settings.provider = 'custom'
  // Cap the per-turn output budget to force the local model's heavy-think turns
  // into max_tokens mid-run (the truncation-continuation rig); absent = default.
  const maxTokens = Number(process.env.GENOFFICE_E2E_MAX_TOKENS ?? '')
  if (Number.isFinite(maxTokens) && maxTokens > 0) settings.maxOutputTokens = maxTokens
  const effort = process.env.GENOFFICE_E2E_EFFORT
  const chatTemplateKwargs =
    effort === 'off' ? { enable_thinking: false } : effort ? { reasoning_effort: effort } : undefined
  // The profile is built, not inlined, so the effort switch and the declared
  // context window can share it without one clobbering the other.
  const profile: AiCustomProfile = {}
  if (chatTemplateKwargs) profile.bodyExtras = { chat_template_kwargs: chatTemplateKwargs }
  if (Number.isFinite(CONTEXT_WINDOW) && CONTEXT_WINDOW > 0)
    profile.modelProfiles = [{ id: MODEL, contextWindow: Math.floor(CONTEXT_WINDOW) }]
  settings.providers.custom = {
    apiKey: 'local',
    model: MODEL,
    baseUrl: BASE_URL,
    ...(Object.keys(profile).length > 0 ? { profile } : {}),
  }
  await writeFile(join(userDataDir, 'ai-settings.json'), JSON.stringify(resolveCustomEndpoints(settings)))

  const summary: Record<string, unknown> = {
    label: LABEL,
    effort: process.env.GENOFFICE_E2E_EFFORT ?? '(endpoint default)',
    maxOutputTokens: settings.maxOutputTokens ?? '(app default)',
    contextWindow:
      Number.isFinite(CONTEXT_WINDOW) && CONTEXT_WINDOW > 0
        ? Math.floor(CONTEXT_WINDOW)
        : '(undeclared — shipped 256KB compaction default)',
    keys: KEYS,
    outCol: OUT_LABEL,
    outsideMode: OUTSIDE,
  }

  const launched = await launchShell({
    userDataDir,
    onboardingSeen: true,
    openFile: workbook,
    videoDir: 'bigtable',
    env: { GENOFFICE_DEBUG_HOOKS: '1', GENOFFICE_AI_TRACE_DIR: traceDir },
  })
  try {
    // 1257×35 through the xlsx sidecar is far slower to open than the toy fixture.
    const sheets = await waitForPageWithUrl(launched.app, '://sheets/', 120_000)
    await sheets.waitForFunction(
      () => document.body.textContent?.includes('Sheet1'),
      null,
      { timeout: 120_000 },
    )
    // The built-app debug hook is exposed only once the workbook finishes
    // loading, which lags the 'Sheet1' label on a slow (multi-MB, cold sidecar)
    // open — on a cold machine the label appeared ~1.5 s before the hook, and
    // the fingerprint loop below read window.__genofficeDebug in that gap.
    // Wait for the hook so cellText never dereferences undefined.
    await sheets.waitForFunction(
      () => (window as unknown as Record<string, unknown>).__genofficeDebug != null,
      null,
      { timeout: 120_000 },
    )

    // ── fingerprint before asking anything ────────────────────────────────────
    // Polled, not read once. The first run of this spec found A1 empty while B1:G1
    // were already correct — on a 3.4 MB sheet that is either the snapshot still
    // hydrating or the import genuinely dropping the first cell, and those lead to
    // completely different conclusions (a flaky test vs a data-loss bug). Polling
    // separates them by time: a header that fills in within seconds is a race, one
    // still empty after a minute is the importer. `headerAttempts` keeps the proof.
    const t0 = Date.now()
    const headerAttempts: Array<{ ms: number; header: string[] }> = []
    let header: string[] = []
    for (;;) {
      header = []
      for (let c = 0; c < EXPECT_HEADER.length; c++) header.push(await cellText(sheets, 0, c))
      headerAttempts.push({ ms: Date.now() - t0, header })
      if (header.every((h, i) => h === EXPECT_HEADER[i])) break
      if (Date.now() - t0 > 60_000) break
      await sheets.waitForTimeout(2_000)
    }
    summary.header = header
    summary.headerAttempts = headerAttempts
    // Evidence for the case this assertion exists to catch: if the model's grid
    // really is missing cells, the screenshot is what tells a rendering problem
    // apart from a model-view problem.
    await mkdir(join(ARTIFACTS_DIR, 'bigtable'), { recursive: true })
    await sheets.screenshot({ path: join(ARTIFACTS_DIR, 'bigtable', `${LABEL}-open.png`) }).catch(() => {})
    expect(header, `not the file the keys came from; header was: ${header.join('|')}`).toEqual(
      EXPECT_HEADER,
    )
    summary.lastRowSeq = await cellText(sheets, EXPECT_ROWS - 1, 0)
    expect(summary.lastRowSeq, `row ${EXPECT_ROWS} should hold the last 序号`).not.toBe('')
    // No probe past the last row on purpose: the sheet's bounds are exactly
    // 1258×35, and getRange throws `Range is out of bounds` rather than returning
    // empty — which turned a fingerprint check into a crash on the first run.
    // The bounds are already proven by that error message, and by the header.

    // Layout fingerprint of the real data, taken before the model touches anything
    // and compared after. The task says "不要改动其他任何单元格", and the first
    // passing run of this spec violated that in force: the sheet has no column past
    // AI (the grid is sized to the data — see MINIMUM_SHEET_COLUMN_COUNT), so
    // writing into AJ was rejected 31 times, and the model's way out was to
    // `insert_cols` twice and then `move_range` 1258 cells back. The four answers
    // were right and the test said so, while the user's table had been re-laid-out.
    // Correct answers are not the whole contract; asserting them alone makes the
    // oracle blind to exactly the damage a "helpful" model does on the way to them.
    const HEADER_COLS = 35
    const layoutOf = async (): Promise<string[]> => {
      const out: string[] = []
      for (let c = 0; c < HEADER_COLS; c++) out.push(await cellText(sheets, 0, c))
      for (let c = 0; c < HEADER_COLS; c++) out.push(await cellText(sheets, EXPECT_ROWS - 1, c))
      return out
    }
    /**
     * The grid's own dimensions, which is the signal the content fingerprint
     * cannot provide. The second run of this spec inserted a column (`Auto-applied
     * 1 change(s): Insert 1 column before column AI` — the bounds error itself
     * suggests that move), wrote its probes into the new column, restored the
     * displaced header text, and put the answers one column further right. Every
     * cell of A..AI ended up byte-identical, so the content fingerprint reported
     * "unchanged" while the sheet had grown from 35 columns to 36. Column count is
     * the one thing that restore cannot hide.
     */
    const dimsOf = (): Promise<{ rows: number; cols: number }> =>
      sheets.evaluate(() => {
        const ws = (window as unknown as DebugWindow).__genofficeDebug.univerAPI
          .getActiveWorkbook()
          ?.getActiveSheet()
        return { rows: ws?.getMaxRows() ?? -1, cols: ws?.getMaxColumns() ?? -1 }
      })
    const layoutBefore = await layoutOf()
    const dimsBefore = await dimsOf()
    summary.layoutBefore = layoutBefore.slice(0, 6)
    summary.dimsBefore = dimsBefore
    const sheetNamesBefore = await sheetNames(sheets)
    summary.sheetNamesBefore = sheetNamesBefore

    const composer = sheets.getByPlaceholder(/Ask AI to analyze/)
    await composer.click()
    await composer.fill(TASK)
    await sheets.getByRole('button', { name: 'Send', exact: true }).click()
    const lines = await waitForRunEnd(traceDir, sheets, 0)

    summary.turns = lines.length
    summary.tools = lines.map((l) => l.response.toolCalls.map((c) => c.name))
    summary.reasoningTotal = lines.reduce((n, l) => n + (l.response.reasoning?.length ?? 0), 0)
    summary.durationMsTotal = lines.reduce((n, l) => n + (l.durationMs ?? 0), 0)
    summary.inputTokensMax = Math.max(0, ...lines.map((l) => l.response.usage?.inputTokens ?? 0))
    // What it actually read: "it guessed" and "it read the wrong range" look
    // identical in the verdict but are different bugs, so keep the inputs.
    summary.readCalls = lines.flatMap((l) =>
      l.response.toolCalls
        .filter((c) => /read|find|context|filter|sort|count/i.test(c.name))
        .map((c) => `${c.name} ${JSON.stringify(c.input ?? {}).slice(0, 200)}`),
    )
    summary.rejected = lines
      .flatMap((l) => (l.messages ?? []).flatMap((m) => (m.results ?? []).map((r) => String(r.output ?? ''))))
      .filter((o) => /Rejected — none of/i.test(o)).length
    summary.staleGate = lines
      .flatMap((l) => (l.messages ?? []).flatMap((m) => (m.results ?? []).map((r) => String(r.output ?? ''))))
      .filter((o) => /structure changed since you last read/i.test(o)).length

    const names = ['conditionalBlank', 'sum', 'over1m', 'year2026'] as const
    const got: Record<string, number | null> = {}
    for (let i = 0; i < names.length; i++)
      got[names[i]!] = asNumber(await cellText(sheets, OUT_ROW + i, OUT_COL))
    summary.got = got
    summary.finalText = lines.at(-1)?.response.text ?? ''

    // Everything is measured before anything is asserted, so a red run still
    // records how the sheet ended up.
    const layoutAfter = await layoutOf()
    const dimsAfter = await dimsOf()
    summary.layoutAfter = layoutAfter.slice(0, 6)
    summary.dimsAfter = dimsAfter
    summary.layoutChanged = layoutBefore.some((v, i) => v !== layoutAfter[i])
    summary.dimsChanged = dimsBefore.rows !== dimsAfter.rows || dimsBefore.cols !== dimsAfter.cols
    // Scratch the model left beside the answers: it wrote probes into AJ6/AJ7
    // while testing whether the cells were writable at all.
    const leftovers: string[] = []
    // Scratch the model leaves behind, scanned across the answer column AND any
    // column the grid grew to cover: P0b made writing past the edge legal, so
    // the model now puts its helper formulas in the first grown column (that is
    // the point), and what matters is that it CLEARS them — the first P0b run
    // wrote 9 helper formulas to AJ2:AJ10 and cleared them in one clear_range.
    // to the edge of the grown grid (capped at +2 past the answer column so the
    // scan stays small): the first P0b run put its helpers at OUT_COL+2 (AJ when
    // the answers were in AH), so stopping at +1 would miss exactly that.
    const grownCols = Math.max(OUT_COL + 1, Math.min(dimsAfter.cols - 1, OUT_COL + 2))
    for (let c = OUT_COL; c <= grownCols; c++) {
      for (let r = OUT_ROW + 4; r < OUT_ROW + 12; r++) {
        const v = await cellText(sheets, r, c)
        if (v.trim() !== '') leftovers.push(`${columnOf(c)}${r + 1}=${v}`)
      }
    }
    summary.leftovers = leftovers

    // Working-sheet mechanism: did the model use a scratch tab, and did it leave
    // one behind? The whole point of the working-sheet guide is that the scratch
    // tab is created, used, and then DELETED — so a clean run ends with only the
    // original sheet(s) present. sheetNamesAfter is the no-leakage proof.
    const sheetNamesAfter = await sheetNames(sheets)
    summary.sheetNamesAfter = sheetNamesAfter
    const workingSheetOps: string[] = []
    for (const line of lines) {
      for (const call of line.response.toolCalls ?? []) {
        if (call.name !== 'propose_operations') continue
        const input = (typeof call.input === 'string' ? safeParse(call.input) : call.input) as
          | { operations?: Array<{ op?: string; name?: string }> }
          | undefined
        for (const op of input?.operations ?? []) {
          if (op.op === 'add_sheet' && op.name === 'GenOffice Scratch') workingSheetOps.push('add')
          if (op.op === 'delete_sheet') workingSheetOps.push('del')
        }
      }
    }
    summary.workingSheetOps = workingSheetOps
    summary.usedWorkingSheet = workingSheetOps.includes('add')

    // Each number is asserted on its own: 3-of-4 means it read *some* of the
    // sheet, which is a different finding from 0-of-4.
    expect(got.conditionalBlank, `${OUT_LABEL}2 expected ${KEYS.conditionalBlank}`).toBe(KEYS.conditionalBlank)
    expect(got.sum, `${OUT_LABEL}3 expected ${KEYS.sum}`).toBe(KEYS.sum)
    expect(got.over1m, `${OUT_LABEL}4 expected ${KEYS.over1m}`).toBe(KEYS.over1m)
    expect(got.year2026, `${OUT_LABEL}5 expected ${KEYS.year2026}`).toBe(KEYS.year2026)

    // Collateral damage. Skipped in OUTSIDE mode on purpose: there the target cell
    // does not exist, so restructuring the sheet is the only way to answer and
    // blaming the model for it would be blaming it for the bug under study.
    if (!OUTSIDE) {
      expect(summary.layoutChanged, 'the task said change nothing else; headers/last row moved').toBe(
        false,
      )
      // The grid may GROW (writing past the edge is legal since P0b — the model
      // uses the first grown column for its helper formulas and must clear them,
      // which the leftovers check below pins), but it must never SHRINK: rows or
      // columns going away means someone's data was deleted.
      expect(
        dimsAfter.rows >= dimsBefore.rows && dimsAfter.cols >= dimsBefore.cols,
        `the grid shrank: ${dimsBefore.rows}x${dimsBefore.cols} -> ${dimsAfter.rows}x${dimsAfter.cols}`,
      ).toBe(true)
      expect(leftovers, 'the task said change nothing else; probe cells left behind').toEqual([])
      // No leaked scratch tab: the working-sheet guide ends by deleting the
      // working sheet, so the workbook must have the same sheet count it started
      // with. A working sheet left behind is exactly the collateral damage the
      // mechanism exists to prevent (and what a crashed/cancelled run would
      // otherwise save into the user's file).
      expect(
        sheetNamesAfter,
        `a scratch working sheet leaked into the workbook (used=${summary.usedWorkingSheet})`,
      ).toEqual(sheetNamesBefore)
    }
  } finally {
    // Written even when an assertion failed — a red run is exactly when the
    // evidence about what the model read matters most.
    summary.traceLines = (await readTrace(traceDir).catch(() => [] as TraceLine[])).length
    await mkdir(join(ARTIFACTS_DIR, 'bigtable'), { recursive: true }).catch(() => {})
    await writeFile(
      join(ARTIFACTS_DIR, `bigtable-${LABEL}.json`),
      JSON.stringify(summary, null, 2),
    ).catch(() => {})
    console.log('BIGTABLE>>>', JSON.stringify(summary, null, 1), '<<<END')
    await launched.app.close().catch(() => {})
  }
})
