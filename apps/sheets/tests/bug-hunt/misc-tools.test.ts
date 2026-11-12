/**
 * 小工具猎捕（~110 题）：补齐 13 个工作簿工具里此前零覆盖的 10 个——
 * read_cells / evaluate_formula / load_guide / read_formats /
 * read_sheet_features / get_workbook_context / select_range /
 * trace_precedents / trace_dependents / create_document。
 *
 * 断言口径与前四个维度一致：走真实 handler（executeWorkbookTool），
 * trace_* 走真实 formula-audit 实现（demo 扫快照 / lazy 走 sidecar stub
 * readWorkbookFormulas），期望字符串凡涉及数据的都从 fixture 推导。
 * 逐字文案 = 模型实际读到的内容，改一个词这里就红。
 */
import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  buildWorkbookContext,
  type ActiveSheetInfo,
  type CreateDocumentToolOutcome,
  type CreateDocumentToolRequest,
  type SheetsSkillDeps,
} from '../../src/renderer/ai/tools'

import {
  COLS,
  demoDeps,
  fixture,
  lazyDeps,
  LAST_COL,
  restoreGlobals,
  ROWS,
  runTool,
  SHEET_ID,
  SHEET_NAME,
} from './harness'

afterEach(() => {
  restoreGlobals()
})

/** fixture 直读（独立 oracle 源），不靠记忆写死 */
const F3_VALUE = String(fixture.cached['F3'])
const F33_VALUE = String(fixture.cached['F33'])
const E3_TEXT = String((fixture.cells['E3'] as { v?: string }).v ?? '')
const A1_TEXT = String((fixture.cells['A1'] as { v?: string }).v ?? '')
const A100_TEXT = String((fixture.cells['A100'] as { v?: string }).v ?? '')

/** 附加一个 sheet 的 demo deps（跨 sheet / 错误值 / 特殊公式场景用） */
function demoWith(
  cells: Record<string, { value?: string | number | boolean | null; formula?: string }>,
) {
  return demoDeps({ extraSheets: [{ id: 's2', name: 'S2', cells }] })
}

function ctxOnly(info: ActiveSheetInfo): string {
  return buildWorkbookContext({ getActiveSheetInfo: () => info } as unknown as SheetsSkillDeps)
}

function lazyInfo(over: Partial<ActiveSheetInfo> = {}): ActiveSheetInfo {
  return {
    mode: 'lazy',
    sheetId: 'sheet-1',
    sheetName: '绝热工程量',
    knownAddresses: [],
    sheets: [{ id: 'sheet-1', name: '绝热工程量', rows: 2313, columns: 7 }],
    ...over,
  } as ActiveSheetInfo
}

// ---------------------------------------------------------------------------
// read_cells
// ---------------------------------------------------------------------------

describe('read_cells', () => {
  it('空数组 → 明确拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_cells', { addresses: [] })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('addresses must be a non-empty array')
  })

  it('非数组 → 同样拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_cells', { addresses: 'A1' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('addresses must be a non-empty array')
  })

  it('未知地址 → (unknown)', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_cells', { addresses: ['Z99'] })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe('Z99: (unknown)')
  })

  it('普通文本 → (text) 后缀', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_cells', { addresses: ['B3'] })
    expect(r.output).toBe('B3: m³ (text)')
  })

  it('数字/日期形状文本 → number/date stored as text 后缀', async () => {
    const b = demoWith({
      A1: { value: '15000.00' },
      A2: { value: '2026-01-15' },
      A3: { value: '1,234,567.89' },
      A4: { value: '50%' },
    })
    const r = await runTool(b.deps, 'read_cells', {
      addresses: ['A1', 'A2', 'A3', 'A4'],
      sheetId: 's2',
    })
    expect(r.isError).toBeFalsy()
    expect(r.output.split('\n')).toEqual([
      'A1: 15000.00 (text: number/date stored as text)',
      'A2: 2026-01-15 (text: number/date stored as text)',
      'A3: 1,234,567.89 (text: number/date stored as text)',
      'A4: 50% (text: number/date stored as text)',
    ])
  })

  it('数值格无后缀', async () => {
    const b = demoWith({ A1: { value: 42 }, A2: { value: true } })
    const r = await runTool(b.deps, 'read_cells', { addresses: ['A1', 'A2'], sheetId: 's2' })
    expect(r.output.split('\n')).toEqual(['A1: 42', 'A2: true'])
  })

  it('demo 公式格（无计算值）→ 只显示公式', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_cells', { addresses: ['F3'] })
    expect(r.output).toBe('F3: =EVALUATE(E3)')
  })

  it('公式格带计算值 → 值 (公式)', async () => {
    const b = demoWith({ A1: { value: 5, formula: '=1+4' } })
    const r = await runTool(b.deps, 'read_cells', { addresses: ['A1'], sheetId: 's2' })
    expect(r.output).toBe('A1: 5 (=1+4)')
  })

  it('超过 100 个地址截断且明示（O4 修复：不再静默）', async () => {
    const b = demoDeps()
    const addresses = Array.from({ length: 150 }, (_, i) => `A${i + 1}`)
    const r = await runTool(b.deps, 'read_cells', { addresses })
    expect(r.isError).toBeFalsy()
    const lines = r.output.split('\n')
    expect(lines.length).toBe(101) // 100 行格值 + 1 行截断提示
    expect(lines[0]).toBe(`A1: ${A1_TEXT} (text)`)
    expect(lines[99]).toBe(`A100: ${A100_TEXT} (text)`)
    // O4 修复：静默截断会让模型以为读全了、对没读到的格下结论
    expect(lines[100]).toBe(
      'Note: only the first 100 of 150 requested addresses were read — the other 50 were not.',
    )
  })

  it('恰好 100 个地址 → 无截断提示', async () => {
    const b = demoDeps()
    const addresses = Array.from({ length: 100 }, (_, i) => `A${i + 1}`)
    const r = await runTool(b.deps, 'read_cells', { addresses })
    expect(r.isError).toBeFalsy()
    expect(r.output.split('\n').length).toBe(100)
    expect(r.output).not.toContain('only the first')
  })

  it('未知 sheetId → 指明从 get_workbook_context 取', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_cells', { addresses: ['A1'], sheetId: 'nope' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('Unknown sheet: nope (use an id from get_workbook_context)')
  })

  it('sheetId 非字符串 → 拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_cells', { addresses: ['A1'], sheetId: 5 })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('sheetId must be a non-empty string when given')
  })

  it('文本格控制字符 → 转义（不撕裂行结构）', async () => {
    const b = demoWith({ A1: { value: 'a\tb' }, A2: { value: 'x\ny' } })
    const r = await runTool(b.deps, 'read_cells', { addresses: ['A1', 'A2'], sheetId: 's2' })
    expect(r.output.split('\n')).toEqual(['A1: a\\tb (text)', 'A2: x\\ny (text)'])
  })

  it('lazy：公式格 → 缓存值 (公式)', async () => {
    const b = lazyDeps()
    const r = await runTool(b.deps, 'read_cells', { addresses: ['F3'] })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe(`F3: ${F3_VALUE} (=EVALUATE(E3))`)
  })

  it('lazy：ensure 失败 → 完整重试提示（含 2000 格拆分指引）', async () => {
    const b = lazyDeps({ ensureRangeLoaded: () => false })
    const r = await runTool(b.deps, 'read_cells', { addresses: ['F3'] })
    expect(r.isError).toBe(true)
    expect(r.output).toBe(
      'The requested cells could not be fully loaded — retry after workbook indexing completes; ' +
        'on streamed workbooks a scatter spanning more than 2000 cells must be split into closer-together read_cells calls.',
    )
  })

  it('lazy：整格超出 extent → 无需流式加载，直接读（ensure 不调用）', async () => {
    const ensure = vi.fn(() => true)
    const b = lazyDeps({ ensureRangeLoaded: ensure })
    const r = await runTool(b.deps, 'read_cells', { addresses: ['B5000'] })
    expect(r.output).toBe('B5000: (unknown)')
    expect(ensure).not.toHaveBeenCalled()
  })

  it('lazy：部分越界 → 收敛到 extent 内再流式加载（截尾不拒）', async () => {
    const ensure = vi.fn(() => true)
    const b = lazyDeps({ ensureRangeLoaded: ensure })
    const r = await runTool(b.deps, 'read_cells', { addresses: ['Z5', 'A99999'] })
    expect(r.isError).toBeFalsy()
    expect(ensure).toHaveBeenCalledTimes(1)
    // sheetId 缺省时是 undefined（parseReadSheetId 不回落 active sheet）
    expect(ensure).toHaveBeenCalledWith(
      { startRow: 4, startColumn: 0, endRow: 2312, endColumn: 6 },
      undefined,
    )
  })

  it('lazy：地址全部无法解析 → 跳过 ensure 直接读', async () => {
    const ensure = vi.fn(() => true)
    const b = lazyDeps({ ensureRangeLoaded: ensure })
    const r = await runTool(b.deps, 'read_cells', { addresses: ['garbage', '!!'] })
    expect(r.output.split('\n')).toEqual(['garbage: (unknown)', '!!: (unknown)'])
    expect(ensure).not.toHaveBeenCalled()
  })
})

// ---------------------------------------------------------------------------
// evaluate_formula
// ---------------------------------------------------------------------------

describe('evaluate_formula', () => {
  it('不以 = 开头 → 拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'evaluate_formula', { formula: '1+1' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('formula must be a string starting with "="')
  })

  it('非字符串 → 拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'evaluate_formula', { formula: 42 })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('formula must be a string starting with "="')
  })

  it('超过 8192 字符 → 拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'evaluate_formula', { formula: '=' + 'a'.repeat(8192) })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('formula is longer than 8192 characters')
  })

  // F3 起 demoDeps() 注入了 demo 求值器，要测 F4 不可用分支须显式去掉
  it('恰好 8192 字符不拒（边界）→ 落到可用性检查', async () => {
    const b = demoDeps()
    const { evaluateFormula: _omit, ...deps } = b.deps
    const r = await runTool(deps, 'evaluate_formula', { formula: '=' + 'a'.repeat(8191) })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('evaluate_formula is not available in this context.')
  })

  it('deps 无 evaluateFormula → 明确不可用 + 回退引导（F4）', async () => {
    const b = demoDeps()
    const { evaluateFormula: _omit, ...deps } = b.deps
    const r = await runTool(deps, 'evaluate_formula', { formula: '=1+1' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('evaluate_formula is not available in this context.')
    expect(r.output).toContain('aggregate_range')
  })

  it('成功 → 值 + 无副作用声明', async () => {
    const b = demoDeps()
    const deps = { ...b.deps, evaluateFormula: vi.fn(async () => 42) }
    const r = await runTool(deps, 'evaluate_formula', { formula: '=1+1' })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe('=1+1 => 42\nNo cell was written; this was evaluated in place.')
  })

  it('结果 null → (empty)', async () => {
    const b = demoDeps()
    const deps = { ...b.deps, evaluateFormula: vi.fn(async () => null) }
    const r = await runTool(deps, 'evaluate_formula', { formula: '=A1' })
    expect(r.output).toBe('=A1 => (empty)\nNo cell was written; this was evaluated in place.')
  })

  it('结果空串 → 同样 (empty)', async () => {
    const b = demoDeps()
    const deps = { ...b.deps, evaluateFormula: vi.fn(async () => '') }
    const r = await runTool(deps, 'evaluate_formula', { formula: '=A1' })
    expect(r.output).toBe('=A1 => (empty)\nNo cell was written; this was evaluated in place.')
  })

  it('结果控制字符 → 转义', async () => {
    const b = demoDeps()
    const deps = { ...b.deps, evaluateFormula: vi.fn(async () => 'a\tb') }
    const r = await runTool(deps, 'evaluate_formula', { formula: '=A1' })
    expect(r.output).toBe('=A1 => a\\tb\nNo cell was written; this was evaluated in place.')
  })

  it('引擎报错 → 原样透出', async () => {
    const b = demoDeps()
    const deps = { ...b.deps, evaluateFormula: vi.fn(async () => ({ error: '#DIV/0!' })) }
    const r = await runTool(deps, 'evaluate_formula', { formula: '=1/0' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('#DIV/0!')
  })

  it('formula 首尾空白被 trim', async () => {
    const b = demoDeps()
    const deps = { ...b.deps, evaluateFormula: vi.fn(async () => 2) }
    const r = await runTool(deps, 'evaluate_formula', { formula: '  =1+1  ' })
    expect(r.output).toBe('=1+1 => 2\nNo cell was written; this was evaluated in place.')
  })

  it('sheetId 透传给引擎', async () => {
    const b = demoWith({ A1: { value: 1 } })
    const evaluate = vi.fn(async () => 1)
    const deps = { ...b.deps, evaluateFormula: evaluate }
    const r = await runTool(deps, 'evaluate_formula', { formula: '=1', sheetId: 's2' })
    expect(r.isError).toBeFalsy()
    expect(evaluate).toHaveBeenCalledWith('=1', 's2')
  })
})

// ---------------------------------------------------------------------------
// load_guide
// ---------------------------------------------------------------------------

describe('load_guide', () => {
  it('空数组 → 拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'load_guide', { guides: [] })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('guides must be a non-empty array')
  })

  it('非数组 → 拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'load_guide', { guides: 'writing' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('guides must be a non-empty array')
  })

  it('未知指南 → 列出全部可用名', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'load_guide', { guides: ['nope'] })
    expect(r.isError).toBe(true)
    expect(r.output).toBe(
      'Unknown guide(s): nope. Available: writing, formatting, layout, structure, charts, data, working-sheet, financial-formatting, data-attribution, table, pivot, shape-image',
    )
  })

  it('未知混已知 → 整体拒绝（不部分加载）', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'load_guide', { guides: ['writing', 'nope'] })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('Unknown guide(s): nope')
  })

  it('单指南 → 正文（无分隔符）', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'load_guide', { guides: ['writing'] })
    expect(r.isError).toBeFalsy()
    expect(r.output.length).toBeGreaterThan(200)
    expect(r.output).not.toContain('\n\n---\n\n')
  })

  it('多指南 → --- 分隔符连接', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'load_guide', { guides: ['writing', 'formatting'] })
    expect(r.isError).toBeFalsy()
    expect(r.output).toContain('\n\n---\n\n')
    const parts = r.output.split('\n\n---\n\n')
    expect(parts.length).toBe(2)
    expect(parts[0]!.length).toBeGreaterThan(100)
    expect(parts[1]!.length).toBeGreaterThan(100)
  })

  it('重复指南名 → 正文只出一遍（O6 修复：去重，纯 token 烧）', async () => {
    const b = demoDeps()
    const single = await runTool(b.deps, 'load_guide', { guides: ['writing'] })
    const dup = await runTool(b.deps, 'load_guide', { guides: ['writing', 'writing'] })
    expect(dup.output).toBe(single.output)
  })

  it('重复跨指南 → 按首次出现顺序各出一遍', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'load_guide', { guides: ['writing', 'formatting', 'writing'] })
    expect(r.isError).toBeFalsy()
    const parts = r.output.split('\n\n---\n\n')
    expect(parts.length).toBe(2)
    const singleWriting = await runTool(b.deps, 'load_guide', { guides: ['writing'] })
    const singleFormatting = await runTool(b.deps, 'load_guide', { guides: ['formatting'] })
    expect(parts[0]).toBe(singleWriting.output)
    expect(parts[1]).toBe(singleFormatting.output)
  })
})

// ---------------------------------------------------------------------------
// read_formats
// ---------------------------------------------------------------------------

describe('read_formats', () => {
  it('空 range → 拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_formats', { range: '  ' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('range must be a non-empty string')
  })

  it('整列跨度 → 点名换写法（含示例）', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_formats', { range: 'A:A' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('whole columns or rows')
    expect(r.output).toContain('A1:A200')
  })

  it('小写地址 → 归一化后接受（大小写不敏感）', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_formats', { range: 'a1:b2' })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe('No explicit formats in this range.')
  })

  it('解析不了的输入 → 给出接受的形状', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_formats', { range: 'a1b2' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('uppercase')
    expect(r.output).not.toContain('whole columns or rows')
  })

  it('超过 200 格 → 拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_formats', { range: 'A1:H26' }) // 208 格
    expect(r.isError).toBe(true)
    expect(r.output).toBe('The range contains more than 200 cells; read it in multiple calls')
  })

  it('恰好 200 格 → 放行', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_formats', { range: 'A1:H25' }) // 200 格
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe('No explicit formats in this range.')
  })

  it('有格式 → 逐格描述（字段顺序固定）', async () => {
    const b = demoDeps()
    const deps = {
      ...b.deps,
      readFormats: () => ({ A1: { bold: true, italic: true, fontSize: 12, wrapText: true } }),
    }
    const r = await runTool(deps, 'read_formats', { range: 'A1' })
    expect(r.output).toBe('A1: bold, italic, size 12, wrap')
  })

  it('范围内无显式格式的格 → (none)；无格式格不出现在结果里', async () => {
    const b = demoDeps()
    const deps = { ...b.deps, readFormats: () => ({ A1: {}, B2: { bold: true } }) }
    const r = await runTool(deps, 'read_formats', { range: 'A1:B2' })
    expect(r.output.split('\n')).toEqual(['A1: (none)', 'B2: bold'])
  })

  it('未知 sheetId → 拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_formats', { range: 'A1', sheetId: 'nope' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('Unknown sheet: nope (use an id from get_workbook_context)')
  })

  it('lazy：ensure 失败 → 索引未完成提示', async () => {
    const b = lazyDeps({ ensureRangeLoaded: () => false })
    const r = await runTool(b.deps, 'read_formats', { range: 'A1' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe(
      'The requested cells could not be fully loaded; retry after workbook indexing completes.',
    )
  })
})

// ---------------------------------------------------------------------------
// read_sheet_features
// ---------------------------------------------------------------------------

describe('read_sheet_features', () => {
  it('demo → 原样透传 feature 文本', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_sheet_features', {})
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe(`Feature state of sheet ${SHEET_NAME} (id=${SHEET_ID}):\nAutoFilter: none`)
  })

  it('未知 sheetId → 拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'read_sheet_features', { sheetId: 'nope' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('Unknown sheet: nope (use an id from get_workbook_context)')
  })

  it('lazy：ensure 失败 → 索引未完成提示', async () => {
    const b = lazyDeps({ ensureRangeLoaded: () => false })
    const r = await runTool(b.deps, 'read_sheet_features', {})
    expect(r.isError).toBe(true)
    expect(r.output).toBe(
      'The requested cells could not be fully loaded; retry after workbook indexing completes.',
    )
  })
})

// ---------------------------------------------------------------------------
// get_workbook_context
// ---------------------------------------------------------------------------

describe('get_workbook_context', () => {
  it('demo 基础行集（4 行，F29 起含 extent 行与数据区行）', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'get_workbook_context', {})
    const lines = r.output.split('\n')
    // F29: demo 替身补齐 rows/columns（真 app demo 分支同款）→ extent 行 +
    // 数据区行出现，与 lazy 基础行集对齐
    expect(lines[0]).toBe(
      `Active sheet: ${SHEET_NAME} (id=${SHEET_ID}, data extent about ${ROWS} rows × ${COLS} columns)`,
    )
    expect(lines[1]).toBe('Mode: demo workbook, current revision=0')
    expect(lines[2]).toBe(
      `Active sheet data area: A1:${LAST_COL}${ROWS} (answer data-size questions directly from this — do not tally block by block with read_range)`,
    )
    // fixture 13266 格 > 200 → O5 修复后的截断标注措辞（总数可见，不再读起来像"只有这些"）
    const total = Object.keys(fixture.cells).length
    expect(total).toBeGreaterThan(200)
    expect(lines[3]).toBe(
      `Known non-empty cells (showing the first 200 of ${total}): ${Object.keys(fixture.cells)
        .slice(0, 200)
        .join(', ')}`,
    )
    expect(lines.length).toBe(4)
  })

  it('lazy 基础行集（4 行，含数据区直答指引）', async () => {
    const b = lazyDeps()
    const r = await runTool(b.deps, 'get_workbook_context', {})
    expect(r.output.split('\n')).toEqual([
      `Active sheet: ${SHEET_NAME} (id=${SHEET_ID}, data extent about 2313 rows × 7 columns)`,
      'Mode: imported real xlsx file (some regions may still be streaming in)',
      'Active sheet data area: A1:G2313 (answer data-size questions directly from this — do not tally block by block with read_range)',
      'No known non-empty cell information yet; read on demand with read_range/read_cells.',
    ])
  })

  it('无工作簿 → 一句话', () => {
    expect(ctxOnly(lazyInfo({ mode: 'none' }))).toBe('No workbook is currently open.')
  })

  it('空 sheet → no data 标注', () => {
    const out = ctxOnly(
      lazyInfo({
        sheetId: 's1',
        sheetName: 'S',
        sheets: [{ id: 's1', name: 'S', rows: 0, columns: 0 }],
      }),
    )
    expect(out).toContain('Active sheet: S (id=s1, no data (empty sheet))')
    expect(out).not.toContain('Active sheet data area:')
  })

  it('流式模式 → 长说明段', () => {
    const out = ctxOnly(lazyInfo({ streaming: true }))
    expect(out).toContain(
      'Mode: imported file in LARGE-FILE STREAMING MODE — the file is too big for a full load, so the grid ',
    )
  })

  it('多 sheet → All sheets 行（含各自 extent）', () => {
    const out = ctxOnly(
      lazyInfo({
        sheets: [
          { id: 'sheet-1', name: '绝热工程量', rows: 2313, columns: 7 },
          { id: 's2', name: '汇总', rows: 10, columns: 3 },
        ],
      }),
    )
    expect(out).toContain(
      `All sheets: ${SHEET_NAME} (id=${SHEET_ID}, data extent about 2313 rows × 7 columns), 汇总 (id=s2, data extent about 10 rows × 3 columns)`,
    )
  })

  it('READ-ONLY 超大 sheet → 警告段', () => {
    const out = ctxOnly(
      lazyInfo({
        sheets: [
          { id: 'sheet-1', name: '绝热工程量', rows: 2313, columns: 7 },
          { id: 's2', name: '汇总', rows: 10, columns: 3, readOnlyOversized: true },
        ],
      }),
    )
    expect(out).toContain(
      'READ-ONLY sheets: 汇总 — the worksheet XML is above the save-patch limit, so edits there can never be saved',
    )
  })

  it('冻结 selection + 整列名 → 语义化指代', () => {
    const out = ctxOnly(
      lazyInfo({
        selection: 'A2:A100',
        selectionFrozen: true,
        selectionColumns: ['数量'],
      }),
    )
    expect(out).toContain(
      'User selection: A2:A100 (the whole "数量" column) — captured when the user sent this message, so it is what "this column / these rows / the selected part" refers to. It stays fixed for the whole run even if the user clicks elsewhere while you work.',
    )
  })

  it('多列 selection → 复数 columns', () => {
    const out = ctxOnly(
      lazyInfo({ selection: 'A2:B100', selectionFrozen: true, selectionColumns: ['a', 'b'] }),
    )
    expect(out).toContain('User selection: A2:B100 (the whole "a", "b" columns) — captured')
  })

  it('实时 selection（未冻结）→ Current selection', () => {
    const out = ctxOnly(lazyInfo({ selection: 'B5' }))
    expect(out).toContain('Current selection: B5')
    expect(out).not.toContain('captured when the user sent this message')
  })

  it('loadedRange → 视口行（注明非数据 extent）', () => {
    const out = ctxOnly(lazyInfo({ loadedRange: 'A1:G500' }))
    expect(out).toContain('Currently loaded viewport: A1:G500 (not the worksheet data extent)')
  })

  it('merges → 合并区行', () => {
    const out = ctxOnly(lazyInfo({ merges: ['A1:C1', 'D2:E2'] }))
    expect(out).toContain('Merged ranges on the active sheet: A1:C1, D2:E2')
  })

  it('charts → 标题行 + 每图一行', () => {
    const out = ctxOnly(
      lazyInfo({
        charts: [{ path: 'charts/0', title: 'T1', types: 'bar', sheetId: 'sheet-1' }],
      }),
    )
    expect(out).toContain(
      'Charts in the workbook (use the path below with edit_chart to edit an existing chart; use add_chart to create one):',
    )
    expect(out).toContain('- charts/0 | title: T1 | type: bar | sheetId: sheet-1')
  })

  it('无标题 chart → (none)', () => {
    const out = ctxOnly(
      lazyInfo({
        charts: [{ path: 'charts/1', title: '', types: 'line', sheetId: 'sheet-1' }],
      }),
    )
    expect(out).toContain('- charts/1 | title: (none) | type: line | sheetId: sheet-1')
  })

  it('knownAddresses 超 200 → 截断且标注总数（O5 修复：不再读起来像"表里只有这些"）', () => {
    const addresses = Array.from({ length: 250 }, (_, i) => `K${i}`)
    const out = ctxOnly(lazyInfo({ knownAddresses: addresses }))
    expect(out).toContain('K199')
    expect(out).not.toContain('K200')
    expect(out).toContain('Known non-empty cells (showing the first 200 of 250):')
  })

  it('knownAddresses ≤ 200 → 维持 (may be incomplete) 措辞', () => {
    const out = ctxOnly(lazyInfo({ knownAddresses: ['A1', 'B2'] }))
    expect(out).toContain('Known non-empty cells (may be incomplete): A1, B2')
    expect(out).not.toContain('showing the first')
  })
})

// ---------------------------------------------------------------------------
// select_range
// ---------------------------------------------------------------------------

describe('select_range', () => {
  it('空串 → 拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'select_range', { range: '' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('range must be a non-empty string')
  })

  it('非字符串 → 拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'select_range', { range: 42 })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('range must be a non-empty string')
  })

  it('整列 → 点名换写法', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'select_range', { range: 'A:A' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('whole columns or rows')
  })

  it('小写 → 归一化后接受（大小写不敏感）', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'select_range', { range: 'a1' })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe(`Selected ${SHEET_NAME}!A1 and scrolled it into view.`)
  })

  it('解析不了的输入 → 给出接受的形状', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'select_range', { range: 'a1b2' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('uppercase')
  })

  it('单格 → 无冒号', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'select_range', { range: 'c10' })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe(`Selected ${SHEET_NAME}!C10 and scrolled it into view.`)
  })

  it('范围 + $ 前缀 → 归一化', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'select_range', { range: '$A$1:$B$2' })
    expect(r.output).toBe(`Selected ${SHEET_NAME}!A1:B2 and scrolled it into view.`)
  })

  it('失败带原因 → 透出原因', async () => {
    const b = demoDeps()
    const deps = { ...b.deps, selectRange: () => ({ ok: false as const, error: 'boom' }) }
    const r = await runTool(deps, 'select_range', { range: 'A1' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('boom')
  })

  it('失败无原因 → 兜底文案', async () => {
    const b = demoDeps()
    const deps = { ...b.deps, selectRange: () => ({ ok: false as const }) }
    const r = await runTool(deps, 'select_range', { range: 'A1' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('Selection failed')
  })

  it('sheetId 透传', async () => {
    const b = demoWith({ A1: { value: 1 } })
    const select = vi.fn(() => ({ ok: true as const, sheetName: 'S2' }))
    const deps = { ...b.deps, selectRange: select }
    const r = await runTool(deps, 'select_range', { range: 'A1', sheetId: 's2' })
    expect(r.isError).toBeFalsy()
    expect(select).toHaveBeenCalledWith('s2', {
      startRow: 0,
      endRow: 0,
      startColumn: 0,
      endColumn: 0,
    })
  })
})

// ---------------------------------------------------------------------------
// trace_precedents（真实 formula-audit 实现）
// ---------------------------------------------------------------------------

describe('trace_precedents', () => {
  it('空地址 → 拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'trace_precedents', { address: '' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('address must be a non-empty string')
  })

  it('范围地址 → 要求单格', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'trace_precedents', { address: 'A1:B2' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('Cannot parse cell address: A1:B2 (expected a single cell like "C10")')
  })

  it('纯数字 → 要求单格', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'trace_precedents', { address: '123' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('Cannot parse cell address: 123 (expected a single cell like "C10")')
  })

  it('非公式格 → 值 + 引导用 trace_dependents', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'trace_precedents', { address: 'A1' })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe(
      `A1 is not a formula cell; value: ${A1_TEXT}. Nothing to trace upstream — use trace_dependents to see what reads it.`,
    )
  })

  it('公式格 → 公式行 + 引用样本行', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'trace_precedents', { address: 'F3' })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe(
      [
        `F3 = ${F3_VALUE} (=EVALUATE(E3))`,
        'Reads 1 reference(s):',
        `- E3 (1 cell(s)): E3=${E3_TEXT}`,
      ].join('\n'),
    )
  })

  it('#VALUE! 公式格（引用空格）→ 错误值进公式行、样本为 (empty)', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'trace_precedents', { address: 'F33' })
    expect(r.output).toBe(
      [
        `F33 = ${F33_VALUE} (=EVALUATE(E33))`,
        'Reads 1 reference(s):',
        '- E33 (1 cell(s)): E33=(empty)',
      ].join('\n'),
    )
  })

  it('引用含错误值 → ⚠️ 标注', async () => {
    const b = demoWith({ A1: { value: '#VALUE!' }, B1: { value: null, formula: '=A1' } })
    const r = await runTool(b.deps, 'trace_precedents', { address: 'B1', sheetId: 's2' })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe(
      [
        'B1 = =A1',
        'Reads 1 reference(s):',
        '- A1 (1 cell(s)) ⚠️ contains error values: A1=#VALUE!',
      ].join('\n'),
    )
  })

  it('跨 sheet 引用（中文表名限定）→ 带表名前缀的样本', async () => {
    const b = demoWith({ B1: { value: null, formula: `=${SHEET_NAME}!A1` } })
    const r = await runTool(b.deps, 'trace_precedents', { address: 'B1', sheetId: 's2' })
    expect(r.output).toBe(
      [
        `B1 = =${SHEET_NAME}!A1`,
        'Reads 1 reference(s):',
        `- ${SHEET_NAME}!A1 (1 cell(s)): A1=${A1_TEXT}`,
      ].join('\n'),
    )
  })

  it('未知表限定 → external 行', async () => {
    const b = demoWith({ B1: { value: null, formula: '=Nope!A1' } })
    const r = await runTool(b.deps, 'trace_precedents', { address: 'B1', sheetId: 's2' })
    expect(r.output).toBe(
      [
        'B1 = =Nope!A1',
        'Reads 1 reference(s):',
        '- Nope!…: external/unresolved reference (another workbook or unknown sheet)',
      ].join('\n'),
    )
  })

  it('defined name → 附 read_sheet_features 指引', async () => {
    const b = demoWith({ B1: { value: null, formula: '=Total+1' } })
    const r = await runTool(b.deps, 'trace_precedents', { address: 'B1', sheetId: 's2' })
    expect(r.output).toBe(
      [
        'B1 = =Total+1',
        'Reads 0 reference(s):',
        'Note: the formula also uses defined names — list them with read_sheet_features and trace their ranges directly.',
      ].join('\n'),
    )
  })

  it('超过 30 个引用 → 30+ 截断 + Note', async () => {
    const formula = '=' + Array.from({ length: 31 }, (_, i) => `A${i + 1}`).join('+')
    const b = demoWith({ B1: { value: null, formula } })
    const r = await runTool(b.deps, 'trace_precedents', { address: 'B1', sheetId: 's2' })
    const lines = r.output.split('\n')
    expect(lines[1]).toBe('Reads 30+ reference(s):')
    expect(lines[lines.length - 1]).toBe('Note: the formula has more references than shown here.')
    expect(lines.length).toBe(1 + 1 + 30 + 1)
  })

  it('未知 sheetId → 拒绝', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'trace_precedents', { address: 'A1', sheetId: 'nope' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('Unknown sheet: nope')
  })

  it('lazy：sidecar 真实路径 → 输出与 demo 一致', async () => {
    const b = lazyDeps()
    const r = await runTool(b.deps, 'trace_precedents', { address: 'F3' })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe(
      [
        `F3 = ${F3_VALUE} (=EVALUATE(E3))`,
        'Reads 1 reference(s):',
        `- E3 (1 cell(s)): E3=${E3_TEXT}`,
      ].join('\n'),
    )
  })

  it('lazy 部分索引：未索引到的格不再误报"非公式"——报"尚未加载"并指引 read_range（O12 修复）', async () => {
    const b = lazyDeps({ sidecar: { indexedThroughRow: 1 } })
    const r = await runTool(b.deps, 'trace_precedents', { address: 'F3' })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe(
      'F3 has not been loaded yet — background indexing has not reached it (or the read failed), ' +
        'so it is unknown whether it holds a formula. Do NOT treat it as a plain value: ' +
        'read_range F3 first (that loads it), then retry trace_precedents.',
    )
  })

  it('lazy 完全索引：确认空格仍报"非公式格"（known-empty 不受 O12 修复影响）', async () => {
    const b = lazyDeps()
    const r = await runTool(b.deps, 'trace_precedents', { address: 'G1' })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe(
      'G1 is not a formula cell; value: (empty). Nothing to trace upstream — use trace_dependents to see what reads it.',
    )
  })

  it('lazy 读失败：F3（fixture 里真有公式）同样不得报"非公式格"——读失败=未知（O12 readFailed 分支）', async () => {
    const b = lazyDeps({ sidecar: { failRead: true } })
    const r = await runTool(b.deps, 'trace_precedents', { address: 'F3' })
    expect(r.isError).toBeFalsy()
    // 若把读失败混同"映射出界=确认空"，这里会输出 "F3 is not a formula cell"——牙试验
    expect(r.output).toBe(
      'F3 has not been loaded yet — background indexing has not reached it (or the read failed), ' +
        'so it is unknown whether it holds a formula. Do NOT treat it as a plain value: ' +
        'read_range F3 first (that loads it), then retry trace_precedents.',
    )
  })
})

// ---------------------------------------------------------------------------
// trace_dependents（真实 formula-audit 实现）
// ---------------------------------------------------------------------------

describe('trace_dependents', () => {
  const NO_DEPS = (addr: string) =>
    `No formulas read ${addr}. (Dependents that reach it only through a defined name are not detected.)`

  it('demo：无引用者', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'trace_dependents', { address: 'A1' })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe(NO_DEPS('A1'))
  })

  it('demo：公式格自身（无人引用 F3）', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'trace_dependents', { address: 'F3' })
    expect(r.output).toBe(NO_DEPS('F3'))
  })

  it('demo：F 列公式引用 E 格', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'trace_dependents', { address: 'E3' })
    expect(r.output).toBe(
      [
        '1 formula cell(s) read E3:',
        `- ${SHEET_NAME}!F3 = ${F3_VALUE} (=EVALUATE(E3))`,
      ].join('\n'),
    )
  })

  it('demo：缓存 #VALUE! 的依赖格原样显示', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'trace_dependents', { address: 'E33' })
    expect(r.output).toBe(
      [
        '1 formula cell(s) read E33:',
        `- ${SHEET_NAME}!F33 = ${F33_VALUE} (=EVALUATE(E33))`,
      ].join('\n'),
    )
  })

  it('跨 sheet 依赖（中文限定名文件公式）→ 主表在前', async () => {
    const b = demoWith({ B1: { value: null, formula: `=${SHEET_NAME}!E3` } })
    const r = await runTool(b.deps, 'trace_dependents', { address: 'E3' })
    expect(r.output).toBe(
      [
        '2 formula cell(s) read E3:',
        `- ${SHEET_NAME}!F3 = ${F3_VALUE} (=EVALUATE(E3))`,
        `- S2!B1 = =${SHEET_NAME}!E3`,
      ].join('\n'),
    )
  })

  it('整列引用覆盖目标格', async () => {
    const b = demoWith({ B1: { value: null, formula: `=${SHEET_NAME}!E:E` } })
    const r = await runTool(b.deps, 'trace_dependents', { address: 'E3' })
    expect(r.output.split('\n')[0]).toBe('2 formula cell(s) read E3:')
    expect(r.output).toContain(`- S2!B1 = =${SHEET_NAME}!E:E`)
  })

  it('超过 100 个依赖 → 100+ 截断 + Note', async () => {
    const cells: Record<string, { value: null; formula: string }> = {}
    for (let i = 1; i <= 105; i += 1) cells[`A${i}`] = { value: null, formula: `=${SHEET_NAME}!E3` }
    const b = demoWith(cells)
    const r = await runTool(b.deps, 'trace_dependents', { address: 'E3' })
    const lines = r.output.split('\n')
    expect(lines[0]).toBe('100+ formula cell(s) read E3:')
    expect(lines[lines.length - 1]).toBe('Note: stopped at the result cap — more dependents exist.')
    expect(lines.length).toBe(1 + 100 + 1)
    expect(lines[lines.length - 2]).toBe(`- S2!A99 = =${SHEET_NAME}!E3`)
  })

  it('lazy：readWorkbookFormulas 全量路径', async () => {
    const b = lazyDeps()
    const r = await runTool(b.deps, 'trace_dependents', { address: 'E3' })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe(
      [
        '1 formula cell(s) read E3:',
        `- ${SHEET_NAME}!F3 = ${F3_VALUE} (=EVALUATE(E3))`,
      ].join('\n'),
    )
  })

  it('lazy 部分索引 → incomplete Note', async () => {
    const b = lazyDeps({ sidecar: { indexedThroughRow: 2 } })
    const r = await runTool(b.deps, 'trace_dependents', { address: 'E3' })
    expect(r.output).toBe(
      [
        '1 formula cell(s) read E3:',
        `- ${SHEET_NAME}!F3 = ${F3_VALUE} (=EVALUATE(E3))`,
        `Note: background indexing has not finished on ${SHEET_NAME} — dependents there may be missing.`,
      ].join('\n'),
    )
  })

  it('lazy 索引未开始 → 0 依赖 + incomplete Note', async () => {
    const b = lazyDeps({ sidecar: { indexedThroughRow: null } })
    const r = await runTool(b.deps, 'trace_dependents', { address: 'E3' })
    expect(r.output).toBe(
      [
        NO_DEPS('E3'),
        `Note: background indexing has not finished on ${SHEET_NAME} — dependents there may be missing.`,
      ].join('\n'),
    )
  })

  it('lazy：journal 会话公式格成为依赖（值回填自引擎）', async () => {
    const b = lazyDeps({ journalCells: [{ row: 0, column: 0, formula: '=E3' }] })
    const r = await runTool(b.deps, 'trace_dependents', { address: 'E3' })
    expect(r.output).toBe(
      [
        '2 formula cell(s) read E3:',
        `- ${SHEET_NAME}!A1 = ${A1_TEXT} (=E3)`,
        `- ${SHEET_NAME}!F3 = ${F3_VALUE} (=EVALUATE(E3))`,
      ].join('\n'),
    )
  })

  it('lazy：journal 静态值覆盖文件公式格 → 影子排除（依赖消失）', async () => {
    const b = lazyDeps({ journalCells: [{ row: 2, column: 5, value: 99 }] })
    const r = await runTool(b.deps, 'trace_dependents', { address: 'E3' })
    expect(r.output).toBe(NO_DEPS('E3'))
  })
})

// ---------------------------------------------------------------------------
// create_document
// ---------------------------------------------------------------------------

describe('create_document', () => {
  const okXlsx: CreateDocumentToolOutcome = {
    ok: true,
    name: '报表.xlsx',
    sheetName: SHEET_NAME,
    hadFormulas: true,
  }
  const withCreate = (outcome: CreateDocumentToolOutcome) => {
    const b = demoDeps()
    const create = vi.fn(async (_request: CreateDocumentToolRequest) => outcome)
    return { b, deps: { ...b.deps, createDocument: create }, create }
  }

  it('deps 无 createDocument → 明确不可用', async () => {
    const b = demoDeps()
    const r = await runTool(b.deps, 'create_document', { type: 'docx', title: 'T', content: 'C' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('create_document is not available in this context.')
  })

  it('type 非法 → 列出全部可选值（可用性检查先于 type 校验）', async () => {
    const { deps } = withCreate(okXlsx)
    const r = await runTool(deps, 'create_document', { type: 'txt' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('type must be one of xlsx/csv/docx/pdf/md/html')
  })

  it('xlsx 带 content → 禁止（指引先写表再导出）', async () => {
    const { b, deps } = withCreate(okXlsx)
    const r = await runTool(deps, 'create_document', { type: 'xlsx', content: '<p>x</p>' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe(
      'content is forbidden for xlsx/csv — they export a worksheet. To create a file from new data, ' +
        'write it into a sheet first (add_sheet + set_range via propose_operations), then export that sheet.',
    )
  })

  it('csv 带 content → 同样禁止', async () => {
    const { b, deps } = withCreate(okXlsx)
    const r = await runTool(deps, 'create_document', { type: 'csv', content: 'a,b' })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('content is forbidden for xlsx/csv')
    void b
  })

  it('xlsx 空白 content（trim 后空）→ 放行到导出', async () => {
    const { deps, create } = withCreate(okXlsx)
    const r = await runTool(deps, 'create_document', { type: 'xlsx', content: '   ' })
    expect(r.isError).toBeFalsy()
    expect(create).toHaveBeenCalledTimes(1)
  })

  it('docx 空 title → 拒绝', async () => {
    const { b, deps } = withCreate(okXlsx)
    const r = await runTool(deps, 'create_document', { type: 'docx', title: ' ', content: 'C' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('title must not be empty')
    void b
  })

  it('docx title 非字符串 → 拒绝', async () => {
    const { b, deps } = withCreate(okXlsx)
    const r = await runTool(deps, 'create_document', { type: 'docx', title: 42, content: 'C' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('title must not be empty')
    void b
  })

  it('docx 空 content → 拒绝', async () => {
    const { b, deps } = withCreate(okXlsx)
    const r = await runTool(deps, 'create_document', { type: 'docx', title: 'T', content: '' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('content must not be empty')
    void b
  })

  it('content 回显 <tool_response> 标签 → 拒绝', async () => {
    const { b, deps } = withCreate(okXlsx)
    const r = await runTool(deps, 'create_document', {
      type: 'docx',
      title: 'T',
      content: '<tool_response>hi</tool_response>',
    })
    expect(r.isError).toBe(true)
    expect(r.output).toBe(
      'content contains a literal <tool_response> tag — that is tool-protocol output, not document content; retry with the actual document HTML',
    )
    void b
  })

  it('回显检查大小写不敏感', async () => {
    const { b, deps } = withCreate(okXlsx)
    const r = await runTool(deps, 'create_document', {
      type: 'pdf',
      title: 'T',
      content: '<TOOL_RESPONSE>x</TOOL_RESPONSE>',
    })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('literal <tool_response> tag')
    void b
  })

  it('content 是裸 JSON dump → 拒绝并给出 HTML 示例', async () => {
    const { b, deps } = withCreate(okXlsx)
    const r = await runTool(deps, 'create_document', {
      type: 'docx',
      title: 'T',
      content: '{"index": 0, "type": "text", "text": "hi"}',
    })
    expect(r.isError).toBe(true)
    expect(r.output).toBe(
      'content contains a raw JSON dump, not an HTML fragment; retry with simple HTML (e.g. <p>…</p>)',
    )
    void b
  })

  it('md 豁免回显检查 → 原样传给 create', async () => {
    const { deps, create } = withCreate({ ok: true, name: 'x.md', path: '/tmp/x.md' })
    const r = await runTool(deps, 'create_document', {
      type: 'md',
      title: 'T',
      content: '<tool_response>x</tool_response>',
    })
    expect(r.isError).toBeFalsy()
    expect(create).toHaveBeenCalledWith({ type: 'md', title: 'T', content: '<tool_response>x</tool_response>' })
  })

  it('xlsx 成功（含公式提示）', async () => {
    const { deps } = withCreate(okXlsx)
    const r = await runTool(deps, 'create_document', { type: 'xlsx' })
    expect(r.isError).toBeFalsy()
    expect(r.output).toBe(
      `Created 报表.xlsx from sheet "${SHEET_NAME}" and opened it in a new tab. The current workbook is unchanged. Note: the sheet contains formulas — the new file holds their current computed values only (no formulas or formatting).`,
    )
  })

  it('xlsx 成功 + path、无公式 → 无提示尾注', async () => {
    const { deps } = withCreate({ ok: true, name: 'x.xlsx', path: '/tmp/x.xlsx', sheetName: 'S' })
    const r = await runTool(deps, 'create_document', { type: 'xlsx' })
    expect(r.output).toBe(
      'Created x.xlsx at /tmp/x.xlsx from sheet "S" and opened it in a new tab. The current workbook is unchanged.',
    )
  })

  it('xlsx 导出失败 → 错误透出', async () => {
    const { b, deps } = withCreate({ ok: false, error: 'export failed' })
    const r = await runTool(deps, 'create_document', { type: 'xlsx' })
    expect(r.isError).toBe(true)
    expect(r.output).toBe('export failed')
    void b
  })

  it('docx 成功 + path → 路径文案', async () => {
    const { deps } = withCreate({ ok: true, name: 'x.docx', path: '/tmp/x.docx' })
    const r = await runTool(deps, 'create_document', { type: 'docx', title: 'T', content: '<p>hi</p>' })
    expect(r.output).toBe('Created the new document at /tmp/x.docx and opened it in a new tab.')
  })

  it('docx 成功无 path → 自保存说明', async () => {
    const { deps } = withCreate({ ok: true, name: 'X' })
    const r = await runTool(deps, 'create_document', { type: 'docx', title: 'T', content: '<p>hi</p>' })
    expect(r.output).toBe('Created the new document "X" in a new tab; it saves itself into the default folder.')
  })

  it('type 缺省 → 按 xlsx 处理', async () => {
    const { deps, create } = withCreate(okXlsx)
    const r = await runTool(deps, 'create_document', {})
    expect(r.isError).toBeFalsy()
    expect(create).toHaveBeenCalledWith({ type: 'xlsx', sheetId: undefined })
  })

  it('docx 不受 sheetId 校验约束', async () => {
    const { deps, create } = withCreate({ ok: true, name: 'x.docx' })
    const r = await runTool(deps, 'create_document', {
      type: 'docx',
      title: 'T',
      content: 'C',
      sheetId: 'nope',
    })
    expect(r.isError).toBeFalsy()
    expect(create).toHaveBeenCalledWith({ type: 'docx', title: 'T', content: 'C' })
  })
})

describe('F15：op 名被当独立工具调 → 错误文案指回 propose_operations（R2 T44）', () => {
  it('edit_chart 当工具发 → 说明它是 op 并给出批次写法', async () => {
    const { deps } = demoDeps({ live: true })
    const r = await runTool(deps, 'edit_chart', { chartPath: 'demo-chart-sheet-1-1', chartType: 'line' })
    expect(r.isError).toBeTruthy()
    expect(r.output).toContain('not a tool')
    expect(r.output).toContain('propose_operations')
    expect(r.output).toContain('{"operations":[{"op":"edit_chart"')
  })

  it('未知工具名 → 列出可用工具 + 提示变更走 propose_operations', async () => {
    const { deps } = demoDeps({ live: true })
    const r = await runTool(deps, 'make_pie', {})
    expect(r.isError).toBeTruthy()
    expect(r.output).toContain('Unknown tool: make_pie')
    expect(r.output).toContain('get_workbook_context')
    expect(r.output).toContain('propose_operations')
  })
})
