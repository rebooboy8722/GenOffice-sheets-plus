/**
 * find_cells 猎捕（~70 题）：
 * - 子串/正则/大小写/errors_only，demo 与 lazy 双路径（lazy plain 走 sidecar stub，
 *   regex/errors_only 走 readSheetRangeMapped 分页扫描——两条都是真实代码）
 * - max_results 钳制边界（0/负数/小数/字符串/199.9/201/恰好命中数）
 * - look_in values/formulas/both 三态（含 sidecar 'formulas'→'formulas_only' 映射契约）
 * - journal 影子（会话写入覆盖文件格 / 会话公式格 fail-open / 打满提前停）
 * - structural ops 后的 file→screen 行号平移（手算断言）
 * - 部分索引 Note 行（sidecar 路径与 range-paged 路径各一）、无 sidecar 降级不崩
 * 命中数/顺序全部由 fixture 独立复算（oracle 复刻 buildMatcher/push 语义，
 * 关键数字已用 python 对 fixture.json 直接复算核对）。
 */
import { afterEach, describe, expect, it } from 'vitest'

import { parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'

import { ERROR_VALUE_RE } from '../../src/renderer/ai/workbook-search'

import {
  demoDeps,
  demoSheetCells,
  engineValue,
  escapeCellText,
  fileCellValue,
  lazyDeps,
  restoreGlobals,
  runTool,
  SHEET_ID,
  SHEET_NAME,
  type LazyDepsOptions,
} from './harness'
import fixture from './fixture.json'

afterEach(() => {
  restoreGlobals()
})

type CellScalar = string | number | null

interface FindOpts {
  query?: string
  regex?: boolean
  look_in?: string
  errors_only?: boolean
  sheetId?: string
  max_results?: number
}

interface Match {
  address: string
  value: CellScalar
  formula?: string
}

interface JournalCell {
  row: number
  column: number
  value?: string | number | null
  formula?: string
}

// ---------------------------------------------------------------------------
// oracle：复刻 tools.ts 的 maxResults 钳制 + buildMatcher + push 截断
// ---------------------------------------------------------------------------

function toolMaxResults(max_results?: unknown): number {
  const n = Number(max_results)
  return Number.isFinite(n) && n >= 1 ? Math.min(Math.floor(n), 200) : 50
}

function buildTest(opts: FindOpts) {
  const look_in = opts.look_in === 'values' || opts.look_in === 'formulas' ? opts.look_in : 'both'
  if (opts.errors_only === true) {
    return (v: CellScalar) => typeof v === 'string' && ERROR_VALUE_RE.test(v)
  }
  const query = typeof opts.query === 'string' ? opts.query.trim() : ''
  let matchText: (text: string) => boolean
  if (opts.regex === true) {
    const pattern = new RegExp(query, 'i')
    matchText = (text) => pattern.test(text)
  } else {
    const needle = query.toLowerCase()
    matchText = (text) => text.toLowerCase().includes(needle)
  }
  const cellValueText = (value: Exclude<CellScalar, null>) =>
    typeof value === 'boolean' ? (value ? '1' : '0') : String(value)
  return (value: CellScalar, formula?: string) => {
    if (look_in !== 'formulas' && value !== null && value !== undefined && matchText(cellValueText(value)))
      return true
    return look_in !== 'values' && !!formula && matchText(formula)
  }
}

/**
 * demo/lazy 共用扫描 oracle（不含 structural 平移——平移用例单独手算断言）。
 * lazy：journal 先行（hasValue 条目），随后按 fixture 行主序（= 真 sidecar
 * chunk 索引序 / demo 快照序）。
 */
function scanMatches(
  mode: 'demo' | 'lazy',
  opts: FindOpts,
  journal: JournalCell[] = [],
): { matches: Match[]; truncated: boolean } {
  const max = toolMaxResults(opts.max_results)
  const test = buildTest(opts)
  const matches: Match[] = []
  let truncated = false
  const push = (m: Match): boolean => {
    if (matches.length >= max) {
      truncated = true
      return false
    }
    matches.push(m)
    return true
  }
  if (mode === 'lazy') {
    const shadowed = new Set(journal.map((j) => `${j.row}:${j.column}`))
    for (const j of journal) {
      const address = `${String.fromCharCode(65 + j.column)}${j.row + 1}`
      let value: CellScalar = j.value ?? null
      if (j.formula && value === null) value = engineValue(address, true) ?? null
      if (test(value, j.formula)) {
        const m: Match = { address, value }
        if (j.formula) m.formula = j.formula
        if (!push(m)) break
      }
    }
    if (truncated) return { matches, truncated }
    for (const [address, entry] of Object.entries(fixture.cells)) {
      if (!entry) continue
      const pos = parseAddress(address)
      if (shadowed.has(`${pos.row}:${pos.column}`)) continue
      const value = fileCellValue(address)
      const f: string | undefined = 'f' in entry ? entry.f : undefined
      if (!test(value, f)) continue
      const m: Match = { address, value }
      if (f) m.formula = f
      if (!push(m)) break
    }
  } else {
    for (const [address, cell] of Object.entries(demoSheetCells())) {
      let value: CellScalar = cell.value
      if (cell.formula && value === null) value = engineValue(address, true) ?? null
      if (!test(value, cell.formula)) continue
      const m: Match = { address, value }
      if (cell.formula) m.formula = cell.formula
      if (!push(m)) break
    }
  }
  return { matches, truncated }
}

function fmtScalar(value: CellScalar, formula?: string): string {
  if (formula) {
    return value === null || value === undefined
      ? escapeCellText(formula)
      : `${escapeCellText(String(value))} (${escapeCellText(formula)})`
  }
  if (value === null) return '(empty)'
  return escapeCellText(String(value))
}

function renderFind(matches: Match[], truncated: boolean, incompleteSheets: string[] = []): string {
  const lines = matches.map((m) => `${SHEET_NAME}!${m.address}: ${fmtScalar(m.value, m.formula)}`)
  const header =
    matches.length === 0
      ? truncated
        ? 'No matching cells found in the scanned region, but the search stopped at the scan budget before covering the whole workbook — do NOT conclude there are no matches; narrow the query or scope with sheetId and retry.'
        : 'No matching cells found.'
      : `${matches.length} matching cell(s)` +
        (truncated
          ? ' (search stopped at the cap — more may exist; narrow the query or scope with sheetId):'
          : ':')
  if (incompleteSheets.length > 0) {
    lines.push(
      `Note: background indexing has not finished on ${incompleteSheets.join(', ')} — matches there may be missing; retry later if something expected is absent.`,
    )
  }
  return [header, ...lines].join('\n')
}

async function expectFind(
  mode: 'demo' | 'lazy',
  opts: FindOpts,
  lazyOpts?: LazyDepsOptions,
): Promise<string> {
  const deps = mode === 'demo' ? demoDeps().deps : lazyDeps(lazyOpts).deps
  const result = await runTool(deps, 'find_cells', opts as Record<string, unknown>)
  expect(result.isError).toBeFalsy()
  const { matches, truncated } = scanMatches(mode, opts, (lazyOpts?.journalCells ?? []) as JournalCell[])
  expect(result.output).toBe(renderFind(matches, truncated))
  return result.output
}

async function expectFindErr(mode: 'demo' | 'lazy', opts: FindOpts, fragment: string): Promise<void> {
  const deps = mode === 'demo' ? demoDeps().deps : lazyDeps().deps
  const result = await runTool(deps, 'find_cells', opts as Record<string, unknown>)
  expect(result.isError).toBe(true)
  expect(result.output).toContain(fragment)
}

// ---------------------------------------------------------------------------
// demo：子串查询
// ---------------------------------------------------------------------------

describe('find_cells demo: 子串', () => {
  it('疑问 → 98 命中（全 G 列，首 G73），默认 50 截断（严格输出）', async () => {
    const out = await expectFind('demo', { query: '疑问' })
    expect(out.startsWith('50 matching cell(s) (search stopped at the cap')).toBe(true)
    expect(out.split('\n')[1]).toBe(`${SHEET_NAME}!G73: ${fmtScalar(fileCellValue('G73'))}`)
  })

  it('疑问 max=98 恰好打完不截断（严格输出）', async () => {
    const out = await expectFind('demo', { query: '疑问', max_results: 98 })
    expect(out.startsWith('98 matching cell(s):')).toBe(true)
  })

  it('疑问 max=99 命中 98 < 99 不截断', async () => {
    const out = await expectFind('demo', { query: '疑问', max_results: 99 })
    expect(out.startsWith('98 matching cell(s):')).toBe(true)
  })

  it.each(['新类型', '半圆法兰盒', '3/4法兰盒', '估算值', '保护层', '0.6mm', '硅酸铝', '单位', 'm³', '项'])(
    '查询 %s 严格输出',
    async (query) => {
      await expectFind('demo', { query, max_results: 200 })
    },
  )

  it('m² → 1217 命中（B/C 列），200 上限截断', async () => {
    const out = await expectFind('demo', { query: 'm²', max_results: 200 })
    expect(out.startsWith('200 matching cell(s) (search stopped at the cap')).toBe(true)
  })

  it('大小写不敏感：M² 与 m² 输出一致', async () => {
    const a = await expectFind('demo', { query: 'M²', max_results: 200 })
    const b = await expectFind('demo', { query: 'm²', max_results: 200 })
    expect(a).toBe(b)
  })

  it('m3（直三）不匹配 m³（上标三）→ 0 命中文案', async () => {
    const deps = demoDeps().deps
    const result = await runTool(deps, 'find_cells', { query: 'm3' })
    expect(result.isError).toBeFalsy()
    expect(result.output).toBe('No matching cells found.')
  })

  it('h=50 → 1652 命中（A/C/D/G 列，首 A3），200 截断', async () => {
    const out = await expectFind('demo', { query: 'h=50', max_results: 200 })
    expect(out.startsWith('200 matching cell(s) (search stopped at the cap')).toBe(true)
    expect(out.split('\n')[1]).toContain('!A3:')
  })

  it('H=60 → 290 命中（A/D 列规格文本"H=60×2层…"，大小写不敏感），200 截断（严格）', async () => {
    const out = await expectFind('demo', { query: 'H=60', max_results: 200 })
    expect(out.startsWith('200 matching cell(s) (search stopped at the cap')).toBe(true)
    expect(out.split('\n')[1]).toContain('!A106:') // python 复算首个命中
  })

  it('无查询且非 errors_only → 拒绝', async () => {
    await expectFindErr('demo', {}, 'query must be a non-empty string')
  })

  it('空白查询 trim 后为空 → 拒绝', async () => {
    await expectFindErr('demo', { query: '   ' }, 'query must be a non-empty string')
  })

  it('query 为数字（非字符串）→ 拒绝', async () => {
    await expectFindErr('demo', { query: 123 as unknown as string }, 'query must be a non-empty string')
  })
})

// ---------------------------------------------------------------------------
// demo：look_in 三态
// ---------------------------------------------------------------------------

describe('find_cells demo: look_in', () => {
  it('evaluate 只在公式里 → look_in=values 0 命中', async () => {
    const deps = demoDeps().deps
    const result = await runTool(deps, 'find_cells', { query: 'evaluate', look_in: 'values' })
    expect(result.isError).toBeFalsy()
    expect(result.output).toBe('No matching cells found.')
  })

  it('look_in=formulas 命中 2310 个 =EVALUATE 公式（全 F 列），200 截断（严格输出）', async () => {
    const out = await expectFind('demo', { query: 'evaluate', look_in: 'formulas', max_results: 200 })
    expect(out.startsWith('200 matching cell(s) (search stopped at the cap')).toBe(true)
    expect(out.split('\n')[1]).toContain('!F3:')
    expect(out.split('\n')[1]).toContain(' (=EVALUATE(E3))')
  })

  it('look_in=both 与 look_in=formulas 对 evaluate 输出一致（值路无命中）', async () => {
    const a = await expectFind('demo', { query: 'evaluate', look_in: 'both', max_results: 200 })
    const b = await expectFind('demo', { query: 'evaluate', look_in: 'formulas', max_results: 200 })
    expect(a).toBe(b)
  })

  it('look_in 非法值按 both 处理（默认）', async () => {
    const a = await expectFind('demo', { query: 'evaluate', look_in: 'bogus', max_results: 200 })
    const b = await expectFind('demo', { query: 'evaluate', max_results: 200 })
    expect(a).toBe(b)
  })

  it('公式格渲染：缓存值 + 公式注记（F3 行严格）', async () => {
    const out = await expectFind('demo', { query: 'evaluate', look_in: 'formulas', max_results: 3 })
    expect(out.split('\n')[1]).toBe(`${SHEET_NAME}!F3: ${fmtScalar(fileCellValue('F3'), '=EVALUATE(E3)')}`)
  })
})

// ---------------------------------------------------------------------------
// demo：正则
// ---------------------------------------------------------------------------

describe('find_cells demo: regex', () => {
  it('regex ^#VALUE!$ → 266 命中，200 截断（严格输出）', async () => {
    const out = await expectFind('demo', { query: '^#VALUE!$', regex: true, max_results: 200 })
    expect(out.startsWith('200 matching cell(s) (search stopped at the cap')).toBe(true)
  })

  it('regex ^#VALUE!$ max=266 被钳到 200（MAX_FIND_RESULTS 硬顶）→ 200 命中 + 截断（严格）', async () => {
    const out = await expectFind('demo', { query: '^#VALUE!$', regex: true, max_results: 266 })
    expect(out.startsWith('200 matching cell(s) (search stopped at the cap')).toBe(true)
  })

  it('regex ^m²$ look_in=values → 200 截断', async () => {
    const out = await expectFind('demo', { query: '^m²$', regex: true, look_in: 'values', max_results: 200 })
    expect(out.startsWith('200 matching cell(s) (search stopped at the cap')).toBe(true)
  })

  it('regex ^=EVALUATE\\(E\\d+\\)$ look_in=formulas → 2310 → 200 截断', async () => {
    const out = await expectFind('demo', {
      query: '^=EVALUATE\\(E\\d+\\)$',
      regex: true,
      look_in: 'formulas',
      max_results: 200,
    })
    expect(out.startsWith('200 matching cell(s) (search stopped at the cap')).toBe(true)
  })

  it('regex 大小写不敏感（i 标志）：^h=50 → 200 截断', async () => {
    const out = await expectFind('demo', { query: '^h=50', regex: true, max_results: 200 })
    expect(out.startsWith('200 matching cell(s) (search stopped at the cap')).toBe(true)
  })

  it.each([['('], ['[abc'], ['*'], ['(?']])('非法正则 %s 被拒（Invalid regex）', async (q: string) => {
    await expectFindErr('demo', { query: q, regex: true }, 'Invalid regex:')
  })

  it('regex 模式下的 0 命中文案', async () => {
    const deps = demoDeps().deps
    const result = await runTool(deps, 'find_cells', { query: '^zzz$', regex: true })
    expect(result.isError).toBeFalsy()
    expect(result.output).toBe('No matching cells found.')
  })
})

// ---------------------------------------------------------------------------
// demo：errors_only
// ---------------------------------------------------------------------------

describe('find_cells demo: errors_only', () => {
  it('默认 max 50 截断，首命中 F33（严格输出）', async () => {
    const out = await expectFind('demo', { errors_only: true })
    expect(out.startsWith('50 matching cell(s) (search stopped at the cap')).toBe(true)
    expect(out.split('\n')[1]).toBe(`${SHEET_NAME}!F33: #VALUE! (=EVALUATE(E33))`)
  })

  it('max=266 被钳到 200 → 200 个 + 截断（266 个错误永远回不全，需收窄）', async () => {
    const out = await expectFind('demo', { errors_only: true, max_results: 266 })
    expect(out.startsWith('200 matching cell(s) (search stopped at the cap')).toBe(true)
  })

  it('max=1 单命中截断（严格输出）', async () => {
    const out = await expectFind('demo', { errors_only: true, max_results: 1 })
    expect(out).toBe(
      `1 matching cell(s) (search stopped at the cap — more may exist; narrow the query or scope with sheetId):\n${SHEET_NAME}!F33: #VALUE! (=EVALUATE(E33))`,
    )
  })

  it('engineHasCache=false → 无计算值 → 0 命中（错误检测依赖引擎值）', async () => {
    const deps = demoDeps({ engineHasCache: false }).deps
    const result = await runTool(deps, 'find_cells', { errors_only: true })
    expect(result.isError).toBeFalsy()
    expect(result.output).toBe('No matching cells found.')
  })
})

// ---------------------------------------------------------------------------
// demo：max_results 钳制边界
// ---------------------------------------------------------------------------

describe('find_cells demo: max_results 钳制', () => {
  it.each([
    [0, 50],
    [-3, 50],
    [199.9, 199],
    ['50', 50],
    [201, 200],
  ] as const)('max_results=%s 生效为 %d', async (raw, effective) => {
    await expectFind('demo', { query: '疑问', max_results: raw as number })
    // 疑问共 98 命中：打满生效上限且仍有剩余才算截断
    const { matches, truncated } = scanMatches('demo', { query: '疑问', max_results: raw as number })
    expect(matches.length).toBe(Math.min(effective, 98))
    expect(truncated).toBe(98 > effective)
  })

  it('未传 max_results → 默认 50', async () => {
    const { matches, truncated } = scanMatches('demo', { query: '疑问' })
    expect(matches.length).toBe(50)
    expect(truncated).toBe(true)
  })
})

// ---------------------------------------------------------------------------
// demo：sheetId
// ---------------------------------------------------------------------------

describe('find_cells demo: sheetId', () => {
  it('unknown sheetId 被拒', async () => {
    await expectFindErr('demo', { query: 'x', sheetId: 'nope' }, 'Unknown sheet: nope')
  })

  it('显式 sheetId 与默认输出一致', async () => {
    const a = await expectFind('demo', { query: '疑问', sheetId: SHEET_ID, max_results: 200 })
    const b = await expectFind('demo', { query: '疑问', max_results: 200 })
    expect(a).toBe(b)
  })

  it('空白 sheetId 按未传处理', async () => {
    const a = await expectFind('demo', { query: '疑问', sheetId: '   ', max_results: 200 })
    const b = await expectFind('demo', { query: '疑问', max_results: 200 })
    expect(a).toBe(b)
  })
})

// ---------------------------------------------------------------------------
// lazy：sidecar plain 路径 + 分页扫描路径
// ---------------------------------------------------------------------------

describe('find_cells lazy', () => {
  it('plain 疑问 max=59 严格输出（sidecar 路径）', async () => {
    const out = await expectFind('lazy', { query: '疑问', max_results: 59 })
    expect(out.startsWith('59 matching cell(s) (search stopped at the cap')).toBe(true)
  })

  it('lazy plain 与 demo 输出一致（同一份数据）', async () => {
    const a = await expectFind('lazy', { query: '疑问', max_results: 59 })
    const b = await expectFind('demo', { query: '疑问', max_results: 59 })
    expect(a).toBe(b)
  })

  it('plain evaluate look_in=formulas 走 sidecar formulas_only，与 demo 一致（严格）', async () => {
    const a = await expectFind('lazy', { query: 'evaluate', look_in: 'formulas', max_results: 200 })
    const b = await expectFind('demo', { query: 'evaluate', look_in: 'formulas', max_results: 200 })
    expect(a).toBe(b)
  })

  it('plain look_in=values 只查值：evaluate 0 命中（验证 values-only 契约）', async () => {
    const deps = lazyDeps().deps
    const result = await runTool(deps, 'find_cells', { query: 'evaluate', look_in: 'values' })
    expect(result.isError).toBeFalsy()
    expect(result.output).toBe('No matching cells found.')
  })

  it('regex 走分页扫描，输出与 demo 一致（严格）', async () => {
    const a = await expectFind('lazy', { query: '^#VALUE!$', regex: true, max_results: 266 })
    const b = await expectFind('demo', { query: '^#VALUE!$', regex: true, max_results: 266 })
    expect(a).toBe(b)
  })

  it('errors_only 默认 50 截断，首命中 F33（严格）', async () => {
    const out = await expectFind('lazy', { errors_only: true })
    expect(out.startsWith('50 matching cell(s) (search stopped at the cap')).toBe(true)
    expect(out.split('\n')[1]).toBe(`${SHEET_NAME}!F33: #VALUE! (=EVALUATE(E33))`)
  })

  it('errors_only max=266 被钳到 200 → 200 命中 + 截断（严格）', async () => {
    const out = await expectFind('lazy', { errors_only: true, max_results: 266 })
    expect(out.startsWith('200 matching cell(s) (search stopped at the cap')).toBe(true)
  })

  it('unknown sheetId 被拒', async () => {
    await expectFindErr('lazy', { query: 'x', sheetId: 'nope' }, 'Unknown sheet: nope')
  })

  it('显式 sheetId 通过', async () => {
    await expectFind('lazy', { query: '疑问', sheetId: SHEET_ID, max_results: 59 })
  })

  it('structural insert-rows(100,1) 后行号平移：文件 F105 → 屏幕 F106（手算）', async () => {
    const deps = lazyDeps({
      structuralOps: [{ kind: 'insert-rows', index: 100, count: 1 }],
    }).deps
    const result = await runTool(deps, 'find_cells', { errors_only: true, max_results: 50 })
    expect(result.isError).toBeFalsy()
    const lines = result.output.split('\n').slice(1)
    // 插入点在 0-based 行 100 之前 → 文件 0-based 行 >100 的命中各 +1
    expect(lines[0]).toContain('!F33:')
    expect(lines[1]).toContain('!F60:')
    expect(lines[2]).toContain('!F106:') // 文件 F105（0-based 104 > 100）→ 屏幕 1-based 106
    expect(result.output).not.toContain('!F105:')
  })
})

// ---------------------------------------------------------------------------
// lazy：journal 影子
// ---------------------------------------------------------------------------

describe('find_cells lazy: journal 影子', () => {
  it('会话写入 ZZZ_MARKER（文件无）→ 仅 journal 命中（严格输出）', async () => {
    const out = await expectFind('lazy', { query: 'ZZZ_MARKER' }, {
      journalCells: [{ row: 5, column: 1, value: 'ZZZ_MARKER' }],
    })
    expect(out).toBe(`1 matching cell(s):\n${SHEET_NAME}!B6: ZZZ_MARKER`)
  })

  it('journal 覆盖 B1 后，文件值 单位 不再命中（B1 是唯一 单位）', async () => {
    const deps = lazyDeps({ journalCells: [{ row: 0, column: 1, value: 'ZZZ' }] }).deps
    const result = await runTool(deps, 'find_cells', { query: '单位' })
    expect(result.isError).toBeFalsy()
    expect(result.output).toBe('No matching cells found.')
  })

  it('journal 公式格（值未计算）fail-open 按公式文本命中（严格输出）', async () => {
    // H1 不在 fixture；文件里 '1+1' 只出现在 E 列普通值格 → look_in=formulas 不查值
    const out = await expectFind(
      'lazy',
      { query: '1+1', look_in: 'formulas' },
      { journalCells: [{ row: 0, column: 7, formula: '=1+1' }] },
    )
    expect(out).toBe(`1 matching cell(s):\n${SHEET_NAME}!H1: =1+1`)
  })

  it('journal 命中数打满 max 时不再扫文件（truncated 提前置位）', async () => {
    const out = await expectFind('lazy', { query: 'J', max_results: 1 }, {
      journalCells: [
        { row: 0, column: 7, value: 'J1' },
        { row: 1, column: 7, value: 'J2' },
      ],
    })
    expect(out).toBe(
      `1 matching cell(s) (search stopped at the cap — more may exist; narrow the query or scope with sheetId):\n${SHEET_NAME}!H1: J1`,
    )
  })
})

// ---------------------------------------------------------------------------
// lazy：索引与降级
// ---------------------------------------------------------------------------

describe('find_cells lazy: 索引与降级', () => {
  it('部分索引（regex 分页路径）→ 命中限索引内（F33/F60）+ 追加 Note 行', async () => {
    const deps = lazyDeps({ sidecar: { indexedThroughRow: 100 } }).deps
    const result = await runTool(deps, 'find_cells', { query: '^#VALUE!$', regex: true, max_results: 5 })
    expect(result.isError).toBeFalsy()
    const lines = result.output.split('\n')
    expect(lines[0]).toBe('2 matching cell(s):')
    expect(lines[1]).toContain('!F33:')
    expect(lines[2]).toContain('!F60:')
    expect(lines[3]).toBe(
      `Note: background indexing has not finished on ${SHEET_NAME} — matches there may be missing; retry later if something expected is absent.`,
    )
  })

  it('部分索引（sidecar plain 路径）→ 命中限索引内 + 追加 Note 行', async () => {
    const deps = lazyDeps({ sidecar: { indexedThroughRow: 100 } }).deps
    const result = await runTool(deps, 'find_cells', { query: '疑问', max_results: 10 })
    expect(result.isError).toBeFalsy()
    const lines = result.output.split('\n')
    expect(lines[0]).toMatch(/^\d+ matching cell\(s\)/)
    expect(result.output).toContain(
      `Note: background indexing has not finished on ${SHEET_NAME} — matches there may be missing; retry later if something expected is absent.`,
    )
  })

  it('无 sidecar（installWindow=false）plain 查询降级不崩，报未完成索引', async () => {
    const deps = lazyDeps({ installWindow: false }).deps
    const result = await runTool(deps, 'find_cells', { query: '疑问' })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('No matching cells found.')
    expect(result.output).toContain(`Note: background indexing has not finished on ${SHEET_NAME}`)
  })
})
