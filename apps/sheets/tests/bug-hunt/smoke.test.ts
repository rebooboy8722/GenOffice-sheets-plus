/**
 * Smoke：先验证 harness 三条路径（demo / lazy / plan）全部打通，再扩 500 题。
 * 每个断言都用 fixture oracle 独立复算值（fmtNum 复刻代码的 en-US 千分位）。
 *
 * 关键行为（dump 实证，非猜测）：
 * - aggregate 数字走 formatNumber→toLocaleString('en-US')：2313 显示 "2,313"。
 * - find_cells 受 MAX_FIND_RESULTS=200 钳制：266 个 #VALUE! 只回 200 + "stopped at the cap"。
 * - read_range 输出不含 sheet 名（只有范围元数据 + TSV + 列 profile）。
 * - lazy 越界：read_range 小范围越界报 "outside the worksheet data extent"；
 *   aggregate_range 有重叠会 clamp、完全无重叠才报 "outside the sheet data extent"。
 * - 孤儿公式行 F 列显示 `#VALUE! (=EVALUATE(E1888))`（缓存值 + 公式注记）。
 * - copy_range 源越界报 "extends beyond the source sheet grid (R rows × C columns)"。
 */
import { afterEach, describe, expect, it } from 'vitest'

import {
  bareTopKey,
  COLS,
  demoDeps,
  fileCellValue,
  fmtNum,
  lazyDeps,
  LAST_COL,
  oracleColumn,
  restoreGlobals,
  ROWS,
  runTool,
  sectionRows,
  SHEET_ID,
} from './harness'

afterEach(() => {
  restoreGlobals()
})

describe('harness smoke: demo mode', () => {
  it('read_range: 表头 + 首数据行 TSV 与 fixture 一致', async () => {
    const { deps } = demoDeps()
    const result = await runTool(deps, 'read_range', { range: 'A1:G3' })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('Read metadata: requested range A1:G3')
    // F29: demo 模式现在也报权威 extent（harness 对齐真 app demo 分支）
    expect(result.output).toContain('authoritative worksheet data extent A1:G2313')
    // 列字母表头行
    expect(result.output).toContain('\tA\tB\tC\tD\tE\tF\tG')
    // 表头行（A1 是普通格）
    expect(result.output).toContain(`1\t${fileCellValue('A1')}`)
    // F3 是公式格：demo 快照 value=null → TSV 输出公式文本
    expect(result.output).toContain('=EVALUATE(E3)')
    // 列 profile 行存在
    expect(result.output).toContain('Column profile (this range)')
  })

  it('read_range: >2000 格先撞 MAX_READ_RANGE_CELLS 上限', async () => {
    const { deps } = demoDeps()
    const result = await runTool(deps, 'read_range', { range: 'A1:G3000' })
    // 3000*7 = 21000 > 2000 → 先撞 2000 格上限
    expect(result.isError).toBe(true)
    expect(result.output).toContain('more than 2000 cells')
  })

  it('aggregate_range: B 列全列统计与 oracle 一致', async () => {
    const { deps } = demoDeps()
    const oracle = oracleColumn('B')
    const result = await runTool(deps, 'aggregate_range', { range: 'B1:B2313' })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain(`cells: ${fmtNum(ROWS)}`)
    expect(result.output).toContain(`non-empty: ${fmtNum(oracle.nonEmpty)}`)
    // B 列全文本（m²/m³/项/单位）
    expect(result.output).toContain('stored as: TEXT')
    const [topKey, topCount] = oracle.top[0]!
    expect(result.output).toContain(`${bareTopKey(topKey)}: ${fmtNum(topCount)}`)
  })

  it('aggregate_range: F 列公式列——2044 数值 + 266 个 #VALUE! 字符串', async () => {
    const { deps } = demoDeps()
    const oracle = oracleColumn('F')
    expect(oracle.formulas).toBe(2310)
    expect(oracle.errorRows.length).toBe(266)
    const result = await runTool(deps, 'aggregate_range', { range: 'F3:F2312' })
    expect(result.isError).toBeFalsy()
    // demo 路径：公式格经 univer 回填缓存值
    expect(result.output).toContain(`non-empty: ${fmtNum(oracle.nonEmpty - 1)}`) // 去掉表头行 F1
    expect(result.output).toContain(`numeric cells: ${fmtNum(oracle.numbers)}`)
    expect(result.output).toContain('#VALUE!')
  })

  it('find_cells: errors_only 受 200 上限钳制，首个错误行命中 oracle', async () => {
    const { deps } = demoDeps()
    const oracle = oracleColumn('F')
    expect(oracle.errorRows.length).toBe(266) // oracle 侧确有 266 个
    const result = await runTool(deps, 'find_cells', { errors_only: true, max_results: 300 })
    expect(result.isError).toBeFalsy()
    // MAX_FIND_RESULTS=200：max_results 300 被钳到 200，266 个只回 200 + 截断提示
    expect(result.output).toContain('200 matching cell(s)')
    expect(result.output).toContain('stopped at the cap')
    // 首个错误行（oracle 复算，1-based）
    expect(result.output).toContain(`!F${oracle.errorRows[0]}`)
  })

  it('find_cells: 普通子串查询命中 E 列文本算式', async () => {
    const { deps } = demoDeps()
    const result = await runTool(deps, 'find_cells', { query: '3.14*', regex: true })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('matching cell(s)')
  })
})

describe('harness smoke: lazy mode', () => {
  it('read_range: 小范围越界被拒（extent 文案）', async () => {
    const { deps } = lazyDeps()
    const result = await runTool(deps, 'read_range', { range: 'A2314:B2314' })
    expect(result.isError).toBe(true)
    expect(result.output).toContain(`outside the worksheet data extent A1:${LAST_COL}${ROWS}`)
  })

  it('read_range: 孤儿公式行 F 列显示缓存错误值 + 公式注记，其余列空', async () => {
    const { deps } = lazyDeps()
    const result = await runTool(deps, 'read_range', { range: 'A1888:G1888' })
    expect(result.isError).toBeFalsy()
    const lines = result.output.split('\n')
    const dataLine = lines.find((l) => l.startsWith('1888\t'))!
    const cells = dataLine.split('\t') // [1888, A, B, C, D, E, F, G]
    // A..E 与 G 全空
    for (const idx of [1, 2, 3, 4, 5, 7]) expect(cells[idx]).toBe('')
    // F 列：缓存 #VALUE! + 公式注记
    expect(cells[6]).toContain('#VALUE!')
    expect(cells[6]).toContain('=EVALUATE(E1888)')
  })

  it('aggregate_range: F 列走 sidecar stub，266 个 #VALUE! 入统计', async () => {
    const { deps } = lazyDeps()
    const oracle = oracleColumn('F')
    const result = await runTool(deps, 'aggregate_range', { range: 'F3:F2312' })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain(`non-empty: ${fmtNum(oracle.nonEmpty - 1)}`)
    expect(result.output).toContain('#VALUE!')
    // 数值统计行：2044 个缓存数值
    expect(result.output).toContain(`numeric cells: ${fmtNum(oracle.numbers)}`)
  })

  it('aggregate_range: 完全越界被拒（extent 文案）', async () => {
    const { deps } = lazyDeps()
    const result = await runTool(deps, 'aggregate_range', { range: 'A3000:G3000' })
    expect(result.isError).toBe(true)
    expect(result.output).toContain('outside the sheet data extent')
  })

  it('aggregate_range: structural ops 后拒绝', async () => {
    const { deps } = lazyDeps({
      structuralOps: [{ kind: 'insert-rows', index: 100, count: 1 }],
    })
    const result = await runTool(deps, 'aggregate_range', { range: 'B1:B10' })
    expect(result.isError).toBe(true)
    expect(result.output).toContain('structural changes')
  })

  it('aggregate_range: 索引未完成被拒', async () => {
    const { deps } = lazyDeps({ sidecar: { indexedThroughRow: 100 } })
    const result = await runTool(deps, 'aggregate_range', { range: 'B1:B2313' })
    expect(result.isError).toBe(true)
    expect(result.output).toContain('still being indexed')
  })

  it('find_cells: plain query 走 sidecar findWorkbookCells', async () => {
    const { deps } = lazyDeps()
    const result = await runTool(deps, 'find_cells', { query: '疑问' })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('matching cell(s)')
  })

  it('find_cells: errors_only 走 readSheetRangeMapped，受 200 上限钳制', async () => {
    const { deps } = lazyDeps()
    const oracle = oracleColumn('F')
    expect(oracle.errorRows.length).toBe(266)
    const result = await runTool(deps, 'find_cells', { errors_only: true, max_results: 300 })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('200 matching cell(s)')
    expect(result.output).toContain(`!F${oracle.errorRows[0]}`)
  })

  it('find_cells: unknown sheetId 被拒', async () => {
    const { deps } = lazyDeps()
    const result = await runTool(deps, 'find_cells', { query: 'x', sheetId: 'nope' })
    expect(result.isError).toBe(true)
    expect(result.output).toContain('Unknown sheet: nope')
  })

  it('propose_operations: set_cell 单格写入通过真实校验链', async () => {
    const { deps } = lazyDeps()
    const result = await runTool(deps, 'propose_operations', {
      operations: [{ op: 'set_cell', sheetId: SHEET_ID, address: 'A2', value: 'test' }],
      summary: 'set A2',
    })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('Auto-applied')
  })

  it('propose_operations: copy_range 源越界被拒（grid 文案）', async () => {
    const { deps } = lazyDeps()
    const result = await runTool(deps, 'propose_operations', {
      operations: [
        {
          op: 'copy_range',
          sheetId: SHEET_ID,
          source: 'A3000:G3000',
          target: 'A1',
        },
      ],
      summary: 'copy',
    })
    expect(result.isError).toBe(true)
    expect(result.output).toContain('extends beyond the source sheet grid')
  })

  it('section oracle: B 列"项"行数 = fixture.sectionRows', () => {
    const oracle = oracleColumn('B')
    const section = sectionRows()
    // "项" 出现次数应等于 section 行数
    const countXiang = oracle.top.find(([v]) => bareTopKey(v) === '项')?.[1] ?? 0
    expect(countXiang).toBe(section.length)
  })
})
