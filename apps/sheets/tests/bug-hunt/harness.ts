/**
 * 500 题 BUG 猎捕基建：fixture.json（openpyxl 直读用户附件生成的独立 oracle）
 * → 三套 deps，全部走**真实** GenOffice 代码路径，只 stub 进程边界（sidecar / Univer 引擎值）：
 *
 * - demoDeps():  demo 模式（内存快照）。read_range / read_cells / aggregate_range /
 *                find_cells 全部真实现；Univer 引擎值用 fixture.cached 回填
 *                （`engineHasCache: false` 变体模拟引擎无缓存的 fail-open 路径）。
 * - lazyDeps():  lazy 模式（真实 xlsx 文件路径）。aggregate_range / find_cells 真实现，
 *                `window.desktopApi.readWorkbookRange` / `findWorkbookCells` 按真实
 *                sidecar schema 从 fixture 造数据（公式格返回缓存 <v>——含 266 个
 *                "#VALUE!" 错误字符串，见 worksheet.rs CellBuilder::finish）。
 * - lazyPlanCtx(): proposeOperations 真实校验链（PlanContext），grid 值来自 fixture。
 *
 * oracle 函数（oracleColumn 等）直接对 fixture 按 GenOffice 显示语义复算，
 * 断言不依赖代码自身输出（独立于被测路径）。
 */
import { vi } from 'vitest'

import { columnLabel, parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { InMemoryWorkbookAdapter } from '@genoffice/xlsx-gateway/domain/in-memory-workbook'
import type { ApplyOutcome, ChangePlan, WorkbookSnapshot } from '@genoffice/xlsx-gateway/domain/workbook.types'
import { aggregateWorkbookRange } from '../../src/renderer/ai/aggregate-range'
import {
  traceWorkbookDependents,
  traceWorkbookPrecedents,
} from '../../src/renderer/ai/formula-audit'
import { findWorkbookCells } from '../../src/renderer/ai/workbook-search'
import { evaluateDemoFormula, type DemoCellMap } from './formula-eval'
import {
  executeWorkbookTool,
  textLooksNumericOrDate,
  type ActiveSheetInfo,
  type FindCellsOptions,
  type SheetsSkillDeps,
  type ToolExecution,
} from '../../src/renderer/ai/tools'
import { ERROR_VALUE_RE } from '../../src/renderer/ai/workbook-search'
import { createEditJournal } from '../../src/renderer/edit-journal'
import { proposeOperations, type PlanContext } from '../../src/renderer/plan-operations'
import type { LazyWorkbookState } from '../../src/renderer/univer-state'
import type { WorkbookReadContext } from '../../src/renderer/ai/workbook-readers'
import fixtureRaw from './fixture.json'

// ---------------------------------------------------------------------------
// fixture 数据
// ---------------------------------------------------------------------------

export interface FixtureCellEntry {
  v?: string | number
  f?: string
}

interface Fixture {
  sheet: { name: string; rows: number; cols: number }
  cells: Record<string, FixtureCellEntry>
  cached: Record<string, string | number>
  colStats: Record<string, { nonEmpty: number; distinct: number; numericText: number }>
  sectionRows: number[]
  functions: string[]
}

export const fixture = fixtureRaw as Fixture

export const SHEET_ID = 'sheet-1'
export const SHEET_NAME = fixture.sheet.name
export const ROWS = fixture.sheet.rows // 2313
export const COLS = fixture.sheet.cols // 7
export const LAST_COL = columnLabel(COLS - 1) // 'G'

/** 文件单元（sidecar 语义）：公式格 → 缓存 <v>（数值或错误字符串），普通格 → 值 */
export function fileCellValue(address: string): string | number | null {
  const entry = fixture.cells[address]
  if (!entry) return null
  if (entry.f) return fixture.cached[address] ?? null
  return entry.v ?? null
}

/** GenOffice 显示值：公式格的显示值来自引擎（= 文件缓存值，除非引擎无缓存） */
export function engineValue(address: string, engineHasCache: boolean): string | number | null {
  const entry = fixture.cells[address]
  if (!entry) return null
  if (entry.f) return engineHasCache ? (fixture.cached[address] ?? null) : null
  return entry.v ?? null
}

export function isFormulaCell(address: string): boolean {
  return (fixture.cells[address]?.f ?? undefined) !== undefined
}

export function formulaText(address: string): string | undefined {
  return fixture.cells[address]?.f
}

// ---------------------------------------------------------------------------
// oracle（独立于被测代码，直接对 fixture 复算 GenOffice 显示语义）
// ---------------------------------------------------------------------------

export interface ColumnOracle {
  nonEmpty: number
  empty: number
  formulas: number
  numbers: number
  texts: number
  numericText: number
  /** 显示值词频 top（数值/字符串混排，键为 `${typeof v}:${v}`） */
  top: [string, number][]
  numberSum: number
  min: number | null
  max: number | null
  /** 显示值为错误字面量的地址（1-based 行号列表） */
  errorRows: number[]
  valueOf: (row1: number) => string | number | null
}

/** 列统计，1-based 行闭区间 [rowLo, rowHi] */
export function oracleColumn(
  letter: string,
  rowLo = 1,
  rowHi = ROWS,
  engineHasCache = true,
): ColumnOracle {
  const column = parseAddress(`${letter}1`).column
  const top = new Map<string, number>()
  let nonEmpty = 0
  let empty = 0
  let formulas = 0
  let numbers = 0
  let texts = 0
  let numericText = 0
  let numberSum = 0
  let min: number | null = null
  let max: number | null = null
  const errorRows: number[] = []
  const valueOf = (row1: number) =>
    engineValue(`${columnLabel(column)}${row1}`, engineHasCache)
  for (let row = rowLo; row <= rowHi; row += 1) {
    const value = valueOf(row)
    if (isFormulaCell(`${columnLabel(column)}${row}`)) formulas += 1
    if (value === null || value === '') {
      empty += 1
      continue
    }
    nonEmpty += 1
    const key = `${typeof value}:${String(value)}`
    top.set(key, (top.get(key) ?? 0) + 1)
    if (typeof value === 'number') {
      numbers += 1
      numberSum += value
      min = min === null ? value : Math.min(min, value)
      max = max === null ? value : Math.max(max, value)
    } else if (typeof value === 'string') {
      texts += 1
      if (textLooksNumericOrDate(value)) numericText += 1
      if (ERROR_VALUE_RE.test(value)) errorRows.push(row)
    }
  }
  const topSorted = [...top.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
  return {
    nonEmpty,
    empty,
    formulas,
    numbers,
    texts,
    numericText,
    top: topSorted,
    numberSum,
    min,
    max,
    errorRows,
    valueOf,
  }
}

/** B 列 = 项 的 section 伪表头行（1-based） */
export function sectionRows(): number[] {
  return fixture.sectionRows
}

/** 复刻 aggregate.ts 的 formatNumber（en-US 千分位）——断言统计行用它拼期望串 */
export function fmtNum(value: number): string {
  return Number.isInteger(value)
    ? value.toLocaleString('en-US')
    : value.toLocaleString('en-US', { maximumFractionDigits: 4 })
}

/** oracle top 键 `${typeof v}:${v}` → 剥离类型前缀，得到 aggregate 输出里的裸值 */
export function bareTopKey(key: string): string {
  const i = key.indexOf(':')
  return i === -1 ? key : key.slice(i + 1)
}

/** 复刻 tools.ts escapeCellText：\ → \\，\t → \t，CR/CRLF → \n，\n → \n（转义形式） */
export function escapeCellText(text: string): string {
  return text
    .replace(/\\/g, '\\\\')
    .replace(/\t/g, '\\t')
    .replace(/\r\n?/g, '\n')
    .replace(/\n/g, '\\n')
}

/**
 * read_range TSV 单格期望值（独立复算 tools.ts L993-1004 的渲染规则，数据只来自 fixture）：
 * - demo：公式格 value=null（快照语义）→ 渲染公式文本；普通格 → 值
 * - lazy：公式格 value=缓存值（engineHasCache 时）→ `value (formula)`；缓存缺失 → 公式文本
 * - 不在 fixture 中的格 → ''（readCells 不返回该地址）
 */
export function renderReadCell(
  address: string,
  mode: 'demo' | 'lazy',
  engineHasCache = true,
): string {
  const entry = fixture.cells[address]
  const formula = entry?.f
  let value: string | number | null
  let exists: boolean
  if (mode === 'demo') {
    value = entry?.f ? null : (entry?.v ?? null)
    exists = entry !== undefined
  } else {
    value = engineValue(address, engineHasCache)
    exists = value !== null || formula !== undefined
  }
  if (!exists) return ''
  if (value === null) return escapeCellText(formula ?? '')
  return formula
    ? `${escapeCellText(String(value))} (${escapeCellText(formula)})`
    : escapeCellText(String(value))
}

/**
 * read_range 数据行 oracle：0-based 闭区间，返回 TSV 行数组（行号 + 各列渲染值）。
 * 与 tools.ts executeRead 的同一段循环逻辑逐格独立复算。
 */
export function expectedReadRows(
  rowLo: number,
  rowHi: number,
  colLo: number,
  colHi: number,
  mode: 'demo' | 'lazy',
  engineHasCache = true,
): string[] {
  const rows: string[] = []
  for (let row = rowLo; row <= rowHi; row += 1) {
    const columns = [String(row + 1)]
    for (let column = colLo; column <= colHi; column += 1) {
      columns.push(renderReadCell(`${columnLabel(column)}${row + 1}`, mode, engineHasCache))
    }
    rows.push(columns.join('\t'))
  }
  return rows
}

/**
 * 列 profile 的每列非空计数 oracle（`L n/m non-empty` 段，kind 不独立断言）。
 * 存在性语义与 readCells 一致：demo = fixture 有该格；lazy = 引擎值非 null 或有公式。
 */
export function expectedProfileParts(
  rowLo: number,
  rowHi: number,
  colLo: number,
  colHi: number,
  mode: 'demo' | 'lazy',
  engineHasCache = true,
): string[] {
  const parts: string[] = []
  for (let column = colLo; column <= colHi; column += 1) {
    const label = columnLabel(column)
    let nonEmpty = 0
    for (let row = rowLo; row <= rowHi; row += 1) {
      const entry = fixture.cells[`${label}${row + 1}`]
      const exists =
        mode === 'demo'
          ? entry !== undefined
          : engineValue(`${label}${row + 1}`, engineHasCache) !== null || entry?.f !== undefined
      if (exists) nonEmpty += 1
    }
    parts.push(`${label} ${nonEmpty}/${rowHi - rowLo + 1} non-empty`)
  }
  return parts
}

export interface RangeAggOracle {
  cells: number
  nonEmpty: number
  distinct: number
  numericCount: number
  sum: number
  min: number | null
  max: number | null
  textCount: number
  numericTextCount: number
  top: [string, number][]
}

/**
 * aggregate_range 统计 oracle：对 0-based 闭区间逐格独立复算（createRangeAggregator 同语义）。
 * 数据值语义与 readRange 一致：demo 公式格 = engineHasCache ? 缓存 : null；lazy 公式格 = 缓存。
 * top 按 (count desc, 首次出现序)——与代码的 Map 插入序一致，便于 top 行严格断言。
 */
export function oracleAggregate(
  rowLo: number,
  rowHi: number,
  colLo: number,
  colHi: number,
  mode: 'demo' | 'lazy',
  engineHasCache = true,
): RangeAggOracle {
  let cells = 0
  let nonEmpty = 0
  let numericCount = 0
  let sum = 0
  let min: number | null = null
  let max: number | null = null
  let textCount = 0
  let numericTextCount = 0
  const counts = new Map<string, number>()
  for (let row = rowLo; row <= rowHi; row += 1) {
    for (let col = colLo; col <= colHi; col += 1) {
      const address = `${columnLabel(col)}${row + 1}`
      const entry = fixture.cells[address]
      let value: string | number | null
      if (!entry) {
        value = null
      } else if (entry.f) {
        value = mode === 'demo' && !engineHasCache ? null : (fixture.cached[address] ?? null)
      } else {
        value = entry.v ?? null
      }
      cells += 1
      if (value === null || value === '') continue
      nonEmpty += 1
      if (typeof value === 'number') {
        numericCount += 1
        sum += value
        min = min === null ? value : Math.min(min, value)
        max = max === null ? value : Math.max(max, value)
      } else if (typeof value === 'string') {
        textCount += 1
        if (textLooksNumericOrDate(value)) numericTextCount += 1
      }
      const key = `${typeof value}:${String(value)}`
      counts.set(key, (counts.get(key) ?? 0) + 1)
    }
  }
  const top = [...counts.entries()].sort((a, b) => b[1] - a[1])
  return {
    cells,
    nonEmpty,
    distinct: counts.size,
    numericCount,
    sum,
    min,
    max,
    textCount,
    numericTextCount,
    top,
  }
}

/** 按 formatRangeAggregate 的分支拼期望输出（top 行单独处理，见调用方） */
export function expectedAggregateLines(o: RangeAggOracle, rangeLabel: string): string[] {
  const lines = [
    `Statistics for ${rangeLabel}:`,
    `- cells: ${fmtNum(o.cells)}, non-empty: ${fmtNum(o.nonEmpty)}`,
    `- distinct values: ${fmtNum(o.distinct)}`,
  ]
  if (o.numericCount > 0) {
    lines.push(
      `- numeric cells: ${fmtNum(o.numericCount)}, sum: ${fmtNum(o.sum)}, ` +
        `average: ${fmtNum(o.numericCount > 0 ? o.sum / o.numericCount : 0)}, min: ${fmtNum(o.min ?? 0)}, max: ${fmtNum(o.max ?? 0)}`,
    )
  }
  if (o.textCount > 0 && o.textCount === o.nonEmpty) {
    lines.push(
      o.numericTextCount === o.textCount
        ? '- stored as: TEXT — every value is numeric/date-shaped but stored as text; plain SUM/COUNT/AVERAGE return 0, use SUMPRODUCT(VALUE(range)) / COUNTA / text criteria'
        : o.numericTextCount > 0
          ? `- stored as: TEXT — ${fmtNum(o.numericTextCount)} of ${fmtNum(o.textCount)} values are numeric/date-shaped text; plain SUM/COUNT/AVERAGE return 0 on those, use SUMPRODUCT(VALUE(range)) / COUNTA / text criteria`
          : '- stored as: TEXT (no numeric cells in this range)',
    )
  } else if (o.textCount > 0 && o.numericCount > 0) {
    lines.push(
      `- stored as: mixed — ${fmtNum(o.numericCount)} numeric, ${fmtNum(o.textCount)} text`,
    )
  }
  return lines
}

/** 只有 F 列公式、其余列全空的行（1-based）——拆分/汇总时最容易踩的脏数据 */
export function orphanFormulaRows(): number[] {
  const rows = new Set<number>()
  for (const [addr, entry] of Object.entries(fixture.cells)) {
    if (!entry.f) continue
    const pos = parseAddress(addr)
    rows.add(pos.row + 1)
  }
  const orphans: number[] = []
  for (const row of rows) {
    let onlyF = true
    for (let c = 0; c < COLS; c += 1) {
      const addr = `${columnLabel(c)}${row}`
      if (c === parseAddress('F1').column) continue
      if (fixture.cells[addr]) {
        onlyF = false
        break
      }
    }
    if (onlyF) orphans.push(row)
  }
  return orphans.sort((a, b) => a - b)
}

// ---------------------------------------------------------------------------
// Univer worksheet / workbook stub
// ---------------------------------------------------------------------------

/** 单格 getRange(string)；多格 getRange(row, col, rows, cols) 两种形态都支持 */
function makeWorksheet(engineHasCache: boolean) {
  const single = (address: string) => ({
    getValue: () => engineValue(address, engineHasCache),
    getRawValue: () => engineValue(address, engineHasCache),
    getRawValues: () => [[engineValue(address, engineHasCache)]],
    getFormula: () => formulaText(address) ?? '',
    getCellDatas: () => [[null]],
  })
  return {
    getSheetId: () => SHEET_ID,
    getSheetName: () => SHEET_NAME,
    getMaxRows: () => ROWS,
    getMaxColumns: () => COLS,
    getRange: (first: unknown, ..._rest: unknown[]) =>
      typeof first === 'string' ? single(first) : { activate: () => undefined },
    scrollToCell: () => undefined,
  }
}

function makeWorkbook(engineHasCache: boolean) {
  const worksheet = makeWorksheet(engineHasCache)
  return {
    getActiveSheet: () => worksheet,
    getSheetBySheetId: (id: string) => (id === SHEET_ID ? worksheet : null),
    getSheets: () => [worksheet],
    setActiveSheet: () => undefined,
  }
}

// ---------------------------------------------------------------------------
// demo 模式（内存快照）
// ---------------------------------------------------------------------------

/** demo 快照 cell 形状：公式格 value=null + formula，普通格 value */
export function demoSheetCells(): Record<string, { value: string | number | null; formula?: string }> {
  const cells: Record<string, { value: string | number | null; formula?: string }> = {}
  for (const [address, entry] of Object.entries(fixture.cells)) {
    cells[address] = entry.f
      ? { value: null, formula: entry.f }
      : { value: entry.v ?? null }
  }
  return cells
}

/**
 * F17/F18: 计数为"非空"（value !== null 或有 formula）的 cell 地址列表——
 * 与 get_workbook_context 的 "Known non-empty cells" 文案一致（见
 * workbook-readers 的 demo 分支）。demo adapter 把被清格原地改写为
 * `{ value: null }` entry 且从不删 key（见
 * packages/xlsx-gateway/tests/in-memory-workbook-clear.test.ts），
 * 裸 Object.keys 计数会虚高，读端必须按内容判空。
 */
export function nonEmptyAddresses<T extends { value: unknown; formula?: unknown }>(
  cells: Record<string, T | undefined>,
): string[] {
  return Object.keys(cells).filter((address) => {
    const cell = cells[address]
    return cell !== undefined && (cell.value !== null || cell.formula !== undefined)
  })
}

/** F29: demo getActiveSheetInfo 必须像真 app（workbook-readers.ts demo 分支
 *  L399-412）一样从 cells 算 extent 报 rows/columns。替身原先只回 {id,name}，
 *  buildWorkbookContext 的 dims() 拿到 undefined 直接吞掉 "data extent about
 *  N rows" 与 "Active sheet data area: A1:G…" 两行——模型无权威行 extent 只能
 *  猜（R13 幻觉主表 2900 行，实 2313：T14 计数对但 extent 错 / T25 删 2801-2900
 *  越界 no-op / T47 报 2800 行）。空表返回 {}（与真 app maxRow<0 分支一致）。 */
export function sheetExtent(
  cells: Record<string, unknown>,
): { rows: number; columns: number } | Record<string, never> {
  let maxRow = -1
  let maxColumn = -1
  for (const address of Object.keys(cells)) {
    const cell = parseAddress(address)
    if (cell.row > maxRow) maxRow = cell.row
    if (cell.column > maxColumn) maxColumn = cell.column
  }
  return maxRow >= 0 ? { rows: maxRow + 1, columns: maxColumn + 1 } : {}
}

export interface DepsBundle {
  ctx: WorkbookReadContext
  deps: SheetsSkillDeps
  state?: LazyWorkbookState
  /** live demo：真实 InMemoryWorkbookAdapter（写闭环独立读回断言用） */
  adapter?: InMemoryWorkbookAdapter
  /** 模拟 Univer 引擎重算结果（进程边界 stub，默认空 = 公式格无计算值） */
  setEngineOverlay?: (values: Record<string, string | number | null>) => void
}

export interface DemoDepsOptions {
  engineHasCache?: boolean
  /** live：真实 InMemoryWorkbookAdapter + 真实 autoApplySafePlan（写闭环） */
  live?: boolean
  /**
   * live 专用：把公式格的缓存值（如 fixture.cached）种进活快照，cell 变为
   * { value: 缓存值, formula } —— 与真实 xlsx"公式 + 上次计算结果"形态一致；
   * 之后对公式格的写入会自然覆盖缓存值（不走 engineOverlay，避免 overlay 遮蔽写回）。
   */
  seedFormulaValues?: Record<string, string | number | null>
  /** autoApplySafePlan 先过此钩子：返回非 null 则直接采用（覆盖 apply 失败文案分支） */
  failApply?: (plan: ChangePlan) => ApplyOutcome | null
  /**
   * live 专用：用既有快照重建 adapter（对话模拟器断点续跑）。恢复后 seedFormulaValues
   * 不再生效（快照里公式格已含缓存值，重种会冲掉运行期写入）。
   */
  initialSnapshot?: WorkbookSnapshot
  /** 附加工作表（多 sheet 场景：跨 sheet 回读前缀 / delete_sheet danglers） */
  extraSheets?: {
    id: string
    name: string
    cells?: Record<string, { value?: string | number | boolean | null; formula?: string }>
  }[]
}

/** 把 adapter.apply 映射成 ApplyOutcome——与 App.tsx autoApplySafePlan 非 lazy 分支同映射 */
export function liveAdapterApply(
  adapter: InMemoryWorkbookAdapter,
  plan: ChangePlan,
): ApplyOutcome {
  try {
    adapter.apply(plan)
    return { ok: true }
  } catch (e) {
    return { ok: false, reason: e instanceof Error ? e.message : String(e) }
  }
}

export function demoDeps(options: DemoDepsOptions = {}): DepsBundle {
  const engineHasCache = options.engineHasCache ?? true
  const live = options.live ?? false
  const restored = options.initialSnapshot?.sheets.find((s) => s.id === SHEET_ID)
  const cells = restored
    ? (restored.cells as Record<string, { value: string | number | boolean | null; formula?: string }>)
    : demoSheetCells()
  if (live && options.seedFormulaValues && !restored) {
    for (const [address, value] of Object.entries(options.seedFormulaValues)) {
      const cell = cells[address]
      if (cell?.formula) cells[address] = { value, formula: cell.formula }
    }
  }
  const engineOverlay = new Map<string, string | number | boolean | null>()
  const sheets =
    options.initialSnapshot?.sheets ??
    [
      { id: SHEET_ID, name: SHEET_NAME, cells },
      ...(options.extraSheets ?? []).map((s) => ({
        id: s.id,
        name: s.name,
        cells: (s.cells ?? {}) as Record<string, { value: string | number | boolean | null; formula?: string }>,
      })),
    ]
  const adapter = live
    ? new InMemoryWorkbookAdapter({ revision: options.initialSnapshot?.revision ?? 0, sheets })
    : undefined
  const ctx = {
    univerRef: { current: { univerAPI: { getActiveWorkbook: () => makeWorkbook(engineHasCache) } } },
    lazyWorkbookRef: { current: null },
    adapterRef: {
      current:
        adapter ?? {
          getSnapshot: () => ({ revision: 0, sheets }),
        },
    },
  } as unknown as WorkbookReadContext
  const info: ActiveSheetInfo = {
    mode: 'demo',
    sheetId: SHEET_ID,
    sheetName: SHEET_NAME,
    revision: 0,
    // F17/F18: 上下文声称 "non-empty"——初始 fixture 里的 null-only entry
    // （合并区锚点格）不计入。
    knownAddresses: nonEmptyAddresses(cells),
    // F29: 带 extent（真 app demo 分支同款），否则 buildWorkbookContext 不报行 extent
    sheets: sheets.map((s) => ({ id: s.id, name: s.name, ...sheetExtent(s.cells) })),
  }
  const deps: SheetsSkillDeps = {
    // live: derive sheets/revision from the live snapshot so add_sheet /
    // delete_sheet show up in get_workbook_context immediately (R1 F2 — the
    // static map left the model unable to reference a sheet it just created).
    // Non-live snapshots are the initial static one, so behavior is unchanged.
    // R3 F12: charts derive from the live snapshot too — the real app's
    // getActiveSheetInfo (workbook-readers demo branch) lists demo visuals,
    // but this hand-rolled stand-in omitted them, so in sim the model could
    // never see a chart it just created and guessed ids until the 8-failure
    // guard stopped the run (R2 T44/T45).
    getActiveSheetInfo: () => {
      if (!adapter) return info
      const snapshot = adapter.getSnapshot()
      // F18: knownAddresses 也必须是 live 的。构造时固化初始 fixture 键表
      // （下方 info 的 Object.keys(cells)）会让 sim 全程恒显 "first 200 of
      // 13266"——R4 T45 模型诚实照抄 stale 数字报"13,266 个非空格"（主表其
      // 实从未被清，真实 12,840）。真 getActiveSheetInfo（workbook-readers
      // demo 分支）每次调用都重算，替身必须跟上。过滤与 F17 读端语义一致
      // （adapter 的 clear_range 会留 {value:null} 键，"non-empty" 只能算
      // 有内容的 entry）。
      const active =
        snapshot.sheets.find((s) => s.id === info.sheetId) ?? snapshot.sheets[0]
      const activeCells = active ? active.cells : {}
      const knownAddresses = Object.keys(activeCells).filter((address) => {
        const cell = activeCells[address]
        return cell !== undefined && (cell.value !== null || cell.formula !== undefined)
      })
      return {
        ...info,
        revision: snapshot.revision,
        // F29: live extent（真 app demo 分支同款）——删行/清列后 extent 随之收缩
        sheets: snapshot.sheets.map((s) => ({ id: s.id, name: s.name, ...sheetExtent(s.cells) })),
        knownAddresses,
        charts: snapshot.sheets.flatMap((entry) =>
          (entry.visuals ?? []).map((visual) => ({
            path: visual.id,
            title: visual.chart?.title ?? '',
            types: visual.chart?.chartTypes.join('+') ?? '',
            sheetId: visual.sheetId,
          })),
        ),
      }
    },
    readCells: (addresses, sheetId) => {
      // live：读活快照（apply 后立即可见）；非 live：初始静态 cells
      const snapshot = live
        ? adapter!.getSnapshot()
        : { revision: 0, sheets }
      const sheet = snapshot.sheets.find((s) => s.id === (sheetId ?? SHEET_ID))
      const sheetCells = sheet?.cells ?? {}
      const out: Record<string, { value: string | number | boolean | null; formula?: string }> = {}
      for (const address of addresses) {
        const cell = sheetCells[address]
        if (!cell) continue
        const value = (engineOverlay.has(address) ? engineOverlay.get(address) : cell.value) ?? null
        out[address] = { value, ...(cell.formula ? { formula: cell.formula } : {}) }
      }
      return out
    },
    readFormats: () => ({}),
    readSheetFeatures: () =>
      `Feature state of sheet ${SHEET_NAME} (id=${SHEET_ID}):\nAutoFilter: none`,
    findCells: (options: FindCellsOptions) => findWorkbookCells(ctx, options),
    selectRange: () => ({ ok: true, sheetName: SHEET_NAME }),
    tracePrecedents: (sheetId, address) => traceWorkbookPrecedents(ctx, sheetId, address),
    traceDependents: (sheetId, address) => traceWorkbookDependents(ctx, sheetId, address),
    aggregateRange: (sheetId, range) => aggregateWorkbookRange(ctx, sheetId, range),
    // F3: demo stand-in for evaluate_formula (no Univer engine here) — a
    // lightweight evaluator over the live snapshot, so newly created sheets
    // are visible to cross-sheet references as well (F2).
    // F19: the sheetId parameter must be honored — production (tools.ts
    // evaluate_formula → parseReadSheetId) defaults a bare ref to the
    // *requested* sheet, not a hardcoded one. The old stand-in always fell
    // back to SHEET_ID, so =COUNTA(A1:F2300) with sheetId=sheet-2 returned
    // the MAIN sheet's number (R5 T10: 13,072 vs aggregate_range's correct
    // 4,675 on sheet-2 — the model burned 215s / 40K reasoning chars
    // reconciling the contradiction). Cross-sheet refs (key set) are
    // unaffected; only the bare-ref fallback follows the requested sheet.
    evaluateFormula: (formula, sheetId) => {
      const snapshot = live ? adapter!.getSnapshot() : { revision: 0, sheets }
      return Promise.resolve(
        evaluateDemoFormula(formula, (key) => {
          const sheet = key
            ? snapshot.sheets.find((s) => s.id === key || s.name === key)
            : snapshot.sheets.find((s) => s.id === (sheetId ?? SHEET_ID))
          return sheet?.cells as DemoCellMap | undefined
        }),
      )
    },
    proposeOperations: (operations, summary) =>
      proposeOperations(
        planCtxFromCtx(
          ctx,
          undefined,
          live
            ? (plan) => {
                const override = options.failApply?.(plan)
                if (override) return override
                return liveAdapterApply(adapter!, plan)
              }
            : undefined,
        ),
        operations,
        summary,
      ),
  }
  const bundle: DepsBundle = { ctx, deps }
  if (live && adapter) {
    bundle.adapter = adapter
    bundle.setEngineOverlay = (values) => {
      engineOverlay.clear()
      for (const [address, value] of Object.entries(values)) {
        engineOverlay.set(address, value)
      }
    }
  }
  return bundle
}

// ---------------------------------------------------------------------------
// lazy 模式（真实文件路径，sidecar stub）
// ---------------------------------------------------------------------------

export interface SidecarOptions {
  /** sidecar 索引到哪一行（0-based，inclusive）；null = 索引未完成 */
  indexedThroughRow?: number | null
  /** 文件隐藏行（0-based） */
  hiddenRows?: number[]
  /** true 时 readWorkbookRange 直接 reject（模拟 sidecar 读失败/会话失效） */
  failRead?: boolean
}

/** 按真实 sidecar schema 造 readWorkbookRange / findWorkbookCells 响应 */
export function sidecarWindow(options: SidecarOptions = {}) {
  const indexedThroughRow =
    options.indexedThroughRow === undefined ? ROWS - 1 : options.indexedThroughRow
  const readWorkbookRange = vi.fn(async (request: {
    range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
  }) => {
    if (options.failRead) throw new Error('sidecar read exploded')
    const range = request.range
    const cells: {
      row: number
      column: number
      value: string | number | null
      formula?: string
    }[] = []
    // 真 sidecar 只返回索引进度内的 cells（readSheetRangeMappedOnce 依据
    // batch.indexedThroughRow 提前 break；"Later batches cannot have data
    // before indexing reaches them"）——部分索引时必须按行裁剪。
    for (const [address, entry] of Object.entries(fixture.cells)) {
      if (!entry) continue
      const pos = parseAddress(address)
      if (
        pos.row >= range.startRow &&
        pos.row <= range.endRow &&
        pos.column >= range.startColumn &&
        pos.column <= range.endColumn
      ) {
        if (indexedThroughRow === null || pos.row > indexedThroughRow) continue
        // 真 sidecar 的 read 结果 cell 带 formula（workbookCellRecordSchema 可选字段；
        // 渲染端 univer-sync.ts:236 `if (cell.formula) range.setFormula(...)` 依赖它
        // 恢复懒表公式）——缺了会让 lazy find 的 range-paged 匹配丢公式注记。
        cells.push({
          row: pos.row,
          column: pos.column,
          value: fileCellValue(address),
          ...(entry.f ? { formula: entry.f } : {}),
        })
      }
    }
    return {
      cells,
      rows: (options.hiddenRows ?? []).map((row) => ({ row, hidden: true })),
      // readSheetRangeMappedOnce 对 merges/hyperlinks 做 push(...batch.x)，
      // 缺字段会 spread undefined 抛错并被误判成"索引未完成"——必须给空数组。
      merges: [] as unknown[],
      hyperlinks: [] as unknown[],
      indexedThroughRow,
      indexingComplete: indexedThroughRow !== null && indexedThroughRow >= ROWS - 1,
    }
  })
  const findWorkbookCells = vi.fn(async (request: {
    query: string
    lookIn?: string
    limit?: number
  }) => {
    const query = request.query.toLowerCase()
    const limit = Math.max(1, request.limit ?? 1000)
    const lookIn = request.lookIn ?? 'all'
    // 与真 sidecar schema 对齐（desktop-api.ts: enum values/formulas/formulas_only/both）：
    // renderer 把工具的 lookIn 'formulas' 映射成 'formulas_only'、'values' 原样传，
    // 所以 'values' 必须按 values-only 解释，'both'/'all' 才是双路。
    const valuesOnly = lookIn === 'values' || lookIn === 'values_only'
    const formulasOnly = lookIn === 'formulas' || lookIn === 'formulas_only'
    const partial = indexedThroughRow !== null && indexedThroughRow < ROWS - 1
    const matches: {
      sheetId: string
      row: number
      column: number
      value: string | number | null
      valueText: string | null
      formulaText?: string
    }[] = []
    for (const [address, entry] of Object.entries(fixture.cells)) {
      if (!entry) continue
      // 部分索引：只能搜到索引进度内的行（真 sidecar 按 chunk 索引扫）。
      if (partial) {
        const pos = parseAddress(address)
        if (pos.row > indexedThroughRow!) continue
      }
      const value = fileCellValue(address)
      const valueText = value === null ? null : String(value)
      const formulaText = entry.f
      let hit = false
      if (!formulasOnly && valueText !== null && valueText.toLowerCase().includes(query)) {
        hit = true
      }
      if (!valuesOnly && formulaText?.toLowerCase().includes(query)) {
        hit = true
      }
      if (!hit) continue
      const pos = parseAddress(address)
      matches.push({
        sheetId: SHEET_ID,
        row: pos.row,
        column: pos.column,
        value,
        valueText,
        ...(formulaText ? { formulaText } : {}),
      })
      if (matches.length >= limit) break
    }
    // 部分索引：complete:false + indexingComplete:false + 无 nextCursor ——
    // renderer findSheetCellsInFile 走 `resumeAt===undefined → complete:false`
    // 即时退出（而非 50ms × 60s 重试循环），工具据此追加 incomplete Note。
    // 观察项：真 sidecar 部分索引时可能自认为"已扫完自己的索引"而报 complete:true，
    // 那样 renderer 不会附 Note，用户拿到无声的局部结果——语义漏洞，留真机核实。
    if (partial) return { matches, complete: false, indexingComplete: false }
    // 完整索引：整表已扫完。limit 打满但还有更多匹配时真 sidecar 应报
    // complete:false + nextCursor 翻页，但 renderer 在 capped 前已由
    // alive()（truncated）退出，工具输出不变，故不实现翻页。
    return { matches, complete: true, indexingComplete: true }
  })
  // readWorkbookFormulas（formula-audit.ts dependents / plan-operations 公式扫描用）：
  // 真 sidecar 返回整表公式格（文件坐标 0-based），值 = 文件缓存 <v>；
  // 部分索引时只有进度内的行，indexingComplete:false → dependents 的
  // incompleteSheets Note 路径。
  const readWorkbookFormulas = vi.fn(async (request: {
    sessionId: string
    sheetId: string
  }) => {
    void request
    // null = 索引还没开始（一格都没有），非 null 且 < 末行 = 部分索引
    const partial = indexedThroughRow === null || indexedThroughRow < ROWS - 1
    const cells: {
      row: number
      column: number
      value: string | number | null
      formula?: string
    }[] = []
    for (const [address, entry] of Object.entries(fixture.cells)) {
      if (!entry?.f) continue
      const pos = parseAddress(address)
      if (partial && (indexedThroughRow === null || pos.row > indexedThroughRow)) continue
      cells.push({
        row: pos.row,
        column: pos.column,
        value: fixture.cached[address] ?? null,
        formula: entry.f,
      })
    }
    cells.sort((a, b) => a.row - b.row || a.column - b.column)
    return {
      cells,
      sharedGroups: [] as unknown[],
      indexingComplete: !partial,
      truncated: false,
    }
  })
  return { readWorkbookRange, findWorkbookCells, readWorkbookFormulas }
}

export interface LazyDepsOptions {
  engineHasCache?: boolean
  /** true 时安装 sidecar window stub（默认 true） */
  installWindow?: boolean
  sidecar?: SidecarOptions
  preloadComplete?: boolean
  formulaMode?: boolean
  /**
   * state.loadedRanges 内容：'full'（默认，覆盖全 grid）或 'none'（空 Map，
   * 用于触发 proposeOperations 的 find/sort 未加载 gate 与 fill 源未加载 gate——
   * 二者都读 loadedRanges + formulaMode / preloadComplete，与 ensureRangeLoaded
   * 回调无关）。
   */
  loadedRanges?: 'full' | 'none'
  ensureRangeLoaded?: (bounds: unknown, sheetId?: string) => boolean | Promise<boolean>
  structuralOps?: { kind: string; index: number; count: number }[]
  /** 预写 journal 单元格（screen 坐标，0-based） */
  journalCells?: { row: number; column: number; value?: string | number | null; formula?: string }[]
}

export function lazyDeps(options: LazyDepsOptions = {}): DepsBundle {
  const engineHasCache = options.engineHasCache ?? true
  const state = {
    file: {
      sessionId: 'session-1',
      sheets: [
        {
          id: SHEET_ID,
          name: SHEET_NAME,
          rowCount: ROWS,
          columnCount: COLS,
          pivotRanges: [],
        },
      ],
      visuals: [],
    },
    generation: 0,
    editJournal: createEditJournal(),
    loadedRanges:
      options.loadedRanges === 'none'
        ? new Map<string, unknown>()
        : new Map<string, unknown>([
            [SHEET_ID, { startRow: 0, endRow: ROWS - 1, startColumn: 0, endColumn: COLS - 1 }],
          ]),
    streamAliases: new Map(),
    loadingKeys: new Map(),
    retryTimers: new Map(),
    sharedFormulaGroups: new Map(),
    pivotDefinitions: new Map(),
    formulaText: new Map(),
    formulaMode: options.formulaMode ?? true,
    flags: { preloadComplete: options.preloadComplete ?? true },
    filterOrigins: new Map(),
    appliedDvSheets: new Set([SHEET_ID]),
  } as unknown as LazyWorkbookState

  if (options.structuralOps?.length) {
    state.editJournal.structuralOps.set(
      SHEET_ID,
      options.structuralOps as unknown as never[],
    )
  }
  for (const cell of options.journalCells ?? []) {
    // createEditJournal() 的 cells 按 sheet 惰性建键，预写前先取/建内层 Map
    let entries = state.editJournal.cells.get(SHEET_ID)
    if (!entries) {
      entries = new Map()
      state.editJournal.cells.set(SHEET_ID, entries)
    }
    entries.set(`${cell.row}:${cell.column}`, {
      row: cell.row,
      column: cell.column,
      hasValue: cell.value !== undefined || cell.formula !== undefined,
      value: cell.value ?? null,
      ...(cell.formula ? { formula: cell.formula } : {}),
    } as never)
  }

  if (options.installWindow ?? true) {
    vi.stubGlobal(
      'window',
      { desktopApi: sidecarWindow(options.sidecar ?? {}) },
    )
  }

  const ctx = {
    univerRef: { current: { univerAPI: { getActiveWorkbook: () => makeWorkbook(engineHasCache) } } },
    lazyWorkbookRef: { current: state },
    adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets: [] }) } },
  } as unknown as WorkbookReadContext
  const info: ActiveSheetInfo = {
    mode: 'lazy',
    sheetId: SHEET_ID,
    sheetName: SHEET_NAME,
    knownAddresses: [],
    sheets: [{ id: SHEET_ID, name: SHEET_NAME, rows: ROWS, columns: COLS }],
  }
  const deps: SheetsSkillDeps = {
    getActiveSheetInfo: () => info,
    readCells: (addresses) => {
      const out: Record<string, { value: string | number | boolean | null; formula?: string }> = {}
      for (const address of addresses) {
        const value = engineValue(address, engineHasCache)
        const formula = formulaText(address)
        if (value === null && formula === undefined) continue
        out[address] = { value, ...(formula ? { formula } : {}) }
      }
      return out
    },
    readFormats: () => ({}),
    readSheetFeatures: () =>
      `Feature state of sheet ${SHEET_NAME} (id=${SHEET_ID}):\nAutoFilter: none`,
    findCells: (findOptions) => findWorkbookCells(ctx, findOptions),
    selectRange: () => ({ ok: true, sheetName: SHEET_NAME }),
    tracePrecedents: (sheetId, address) => traceWorkbookPrecedents(ctx, sheetId, address),
    traceDependents: (sheetId, address) => traceWorkbookDependents(ctx, sheetId, address),
    ensureRangeLoaded:
      options.ensureRangeLoaded ??
      ((bounds) => {
        const b = bounds as { startRow: number; endRow: number; startColumn: number; endColumn: number }
        return (
          b.startRow >= 0 &&
          b.startColumn >= 0 &&
          b.endRow < ROWS &&
          b.endColumn < COLS
        )
      }),
    aggregateRange: (sheetId, range) => aggregateWorkbookRange(ctx, sheetId, range),
    proposeOperations: (operations, summary) =>
      proposeOperations(planCtxFromCtx(ctx, state), operations, summary),
  }
  return { ctx, deps, state }
}

// ---------------------------------------------------------------------------
// PlanContext（proposeOperations 真实校验链）
// ---------------------------------------------------------------------------

function planCtxFromCtx(
  ctx: WorkbookReadContext,
  state?: LazyWorkbookState,
  autoApply?: (plan: ChangePlan) => ApplyOutcome | Promise<ApplyOutcome>,
): PlanContext {
  return {
    adapterRef: ctx.adapterRef,
    univerRef: ctx.univerRef,
    lazyWorkbookRef: { current: state ?? ctx.lazyWorkbookRef.current },
    lazyPreviewRef: { current: null },
    setPreview: vi.fn(),
    // live demo 传真 apply（liveAdapterApply）；其余路径保持 stub {ok:true}
    // （lazy 写闭环走 handleLazyApply，不在 demo 写闭环范围）
    autoApplySafePlan: vi.fn((plan: ChangePlan) =>
      autoApply ? Promise.resolve(autoApply(plan)) : Promise.resolve({ ok: true }),
    ),
  } as unknown as PlanContext
}

export function restoreGlobals(): void {
  vi.unstubAllGlobals()
}

// ---------------------------------------------------------------------------
// 工具调用 helper
// ---------------------------------------------------------------------------

/** 统一工具入口：AgentToolCall → ToolExecution（同步/异步都归一成 Promise） */
export function runTool(
  deps: SheetsSkillDeps,
  name: string,
  input: Record<string, unknown> = {},
): Promise<ToolExecution> {
  const result = executeWorkbookTool({ id: 't', name, input }, deps)
  return Promise.resolve(result)
}
