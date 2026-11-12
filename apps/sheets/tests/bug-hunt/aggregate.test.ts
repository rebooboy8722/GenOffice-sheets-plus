/**
 * aggregate_range 猎捕（~270 题）：
 * - 全列 A-G × 全段/表头段/中段/表尾段 × demo/lazy，输出逐行严格断言（oracle 独立复算）
 * - 105 个 section 段（B 列"项"伪表头切分）× demo/lazy —— 用户真实需求"按段汇总"
 * - F 列 2044 缓存数值 + 266 个 #VALUE!（mixed 分支）、孤儿行段
 * - topValues 参数（默认 10 / 3 / 50 / 0 / 负数钳制）
 * - lazy 专属：extent 拒绝、structural 拒绝、索引未完成/部分索引
 * - engineHasCache=false：demo F 列退化为全空
 */
import { afterEach, describe, expect, it } from 'vitest'

import { parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'

import {
  bareTopKey,
  demoDeps,
  expectedAggregateLines,
  fmtNum,
  lazyDeps,
  LAST_COL,
  orphanFormulaRows,
  oracleAggregate,
  restoreGlobals,
  ROWS,
  runTool,
  sectionRows,
  SHEET_ID,
  type DepsBundle,
  type RangeAggOracle,
} from './harness'

afterEach(() => {
  restoreGlobals()
})

function makeDeps(mode: 'demo' | 'lazy', engineHasCache = true): DepsBundle {
  return mode === 'demo' ? demoDeps({ engineHasCache }) : lazyDeps({ engineHasCache })
}

function topLine(value: string, count: number): string {
  const shown = value.length > 80 ? `${value.slice(0, 80)}…` : value
  return `  ${shown}: ${fmtNum(count)}`
}

/** 严格断言整个 aggregate 输出（统计行 + top 行） */
async function expectAggregate(
  mode: 'demo' | 'lazy',
  raw: string,
  engineHasCache = true,
  topValues?: number,
): Promise<void> {
  const { deps } = makeDeps(mode, engineHasCache)
  const input: Record<string, unknown> = { range: raw }
  if (topValues !== undefined) input.topValues = topValues
  const result = await runTool(deps, 'aggregate_range', input)
  expect(result.isError).toBeFalsy()
  const bounds = parseRange(raw.trim().toUpperCase())
  const o = oracleAggregate(
    bounds.startRow,
    bounds.endRow,
    bounds.startColumn,
    bounds.endColumn,
    mode,
    engineHasCache,
  )
  const nTop =
    topValues === undefined
      ? 10
      : Number.isFinite(topValues)
        ? Math.min(Math.max(Math.floor(topValues), 0), 50)
        : 10
  const expected = [...expectedAggregateLines(o, raw.trim().toUpperCase())]
  const topN = o.top.slice(0, nTop)
  if (topN.length > 0) {
    expected.push('- top values by frequency:')
    for (const [key, count] of topN) expected.push(topLine(bareTopKey(key), count))
  }
  expect(result.output).toBe(expected.join('\n'))
}

async function expectAggregateErr(
  mode: 'demo' | 'lazy',
  raw: string,
  fragment: string,
  options?: Parameters<typeof lazyDeps>[0],
): Promise<void> {
  const { deps } = makeDeps(mode, options?.engineHasCache)
  if (mode === 'lazy' && options) {
    // 重新带完整 options 建 deps（structuralOps / sidecar 等）
    const rebuilt = lazyDeps(options)
    const result = await runTool(rebuilt.deps, 'aggregate_range', { range: raw })
    expect(result.isError).toBe(true)
    expect(result.output).toContain(fragment)
    return
  }
  const result = await runTool(deps, 'aggregate_range', { range: raw })
  expect(result.isError).toBe(true)
  expect(result.output).toContain(fragment)
}

// ---------------------------------------------------------------------------
// 全列 × 典型段
// ---------------------------------------------------------------------------

describe('aggregate demo: 全列与典型段', () => {
  it.each(['A1:A2313', 'B1:B2313', 'C1:C2313', 'D1:D2313', 'E1:E2313', 'F1:F2313', 'G1:G2313'])(
    'demo 全列 %s 输出与 oracle 一致',
    async (range) => {
      await expectAggregate('demo', range)
    },
  )

  it.each([
    'F3:F2312', // 去掉表头
    'A1:A50',
    'B1000:B1050',
    'C2260:C2313',
    'E1:E20',
    'G1:G100',
    'A1:G20', // 矩形：demo 公式格无值 → 全 TEXT
    'B1:F2313',
  ])('demo 段 %s 输出与 oracle 一致', async (range) => {
    await expectAggregate('demo', range)
  })

  it('demo: B 列 stored as TEXT 分支 + top 值 m²/m³/项/单位', async () => {
    const { deps } = demoDeps()
    const result = await runTool(deps, 'aggregate_range', { range: 'B1:B2313' })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('- stored as: TEXT (no numeric cells in this range)')
    expect(result.output).toContain(`  m²: ${fmtNum(1210)}`)
    expect(result.output).toContain(`  m³: ${fmtNum(989)}`)
  })
})

describe('aggregate lazy: 全列与典型段', () => {
  it.each(['A1:A2313', 'B1:B2313', 'C1:C2313', 'D1:D2313', 'E1:E2313', 'F1:F2313', 'G1:G2313'])(
    'lazy 全列 %s 输出与 oracle 一致',
    async (range) => {
      await expectAggregate('lazy', range)
    },
  )

  it.each([
    'F3:F2312',
    'A1:A50',
    'B1000:B1050',
    'C2260:C2313',
    'E1:E20',
    'G1:G100',
    'A1:G20', // 矩形：lazy F 列有缓存数值 → mixed
    'B1:F2313',
    `A1:${LAST_COL}2313`, // 整表 16191 格
  ])('lazy 段 %s 输出与 oracle 一致', async (range) => {
    await expectAggregate('lazy', range)
  })

  it('lazy: F 列 mixed 分支——2044 数值 + 266 个 #VALUE! 文本', async () => {
    const { deps } = lazyDeps()
    const result = await runTool(deps, 'aggregate_range', { range: 'F3:F2312' })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('- stored as: mixed — 2,044 numeric, 266 text')
    expect(result.output).toContain('numeric cells: 2,044, sum: 4,225.2303')
    expect(result.output).toContain('min: 0.0003')
    expect(result.output).toContain('max: 61.0727')
  })
})

// ---------------------------------------------------------------------------
// 孤儿公式行（只有 F 列有 =EVALUATE(...)）
// ---------------------------------------------------------------------------

describe('aggregate 孤儿公式行', () => {
  it('fixture 确有 8 个孤儿行', () => {
    expect(orphanFormulaRows()).toEqual([1888, 1889, 1915, 1916, 1934, 1935, 1967, 1968])
  })

  it.each([['demo'], ['lazy']] as const)('孤儿行 %s 单行聚合（%s）', async (mode: string) => {
    for (const row of orphanFormulaRows()) {
      await expectAggregate(mode as 'demo' | 'lazy', `A${row}:G${row}`)
    }
  })

  it.each([['demo'], ['lazy']] as const)('孤儿行覆盖段 F1885:F1970（%s）', async (mode: string) => {
    await expectAggregate(mode as 'demo' | 'lazy', 'F1885:F1970')
  })
})

// ---------------------------------------------------------------------------
// section 段（B 列"项"伪表头切分，105 段）——用户真实需求"按段汇总"
// ---------------------------------------------------------------------------

describe('aggregate section 段（B 列"项"切分）', () => {
  const sections: [number, number, number][] = []
  const marks = sectionRows()
  for (let i = 0; i < marks.length; i += 1) {
    const start = marks[i]!
    const end = i + 1 < marks.length ? marks[i + 1]! - 1 : ROWS
    sections.push([i, start, end])
  }

  it('section 切分覆盖整表', () => {
    expect(marks[0]).toBeGreaterThanOrEqual(1)
    for (let i = 1; i < marks.length; i += 1) expect(marks[i]).toBeGreaterThan(marks[i - 1]!)
    expect(sections[sections.length - 1]![2]).toBe(ROWS)
  })

  it.each(sections)('demo 段 #%d 行 %d–%d 聚合与 oracle 一致', async (_i, start, end) => {
    await expectAggregate('demo', `A${start}:G${end}`)
  })

  it.each(sections)('lazy 段 #%d 行 %d–%d 聚合与 oracle 一致', async (_i, start, end) => {
    await expectAggregate('lazy', `A${start}:G${end}`)
  })

  it.each(sections)('lazy 段 #%d F 列聚合与 oracle 一致', async (_i, start, end) => {
    await expectAggregate('lazy', `F${start}:F${end}`)
  })
})

// ---------------------------------------------------------------------------
// topValues 参数
// ---------------------------------------------------------------------------

describe('aggregate topValues 参数', () => {
  it.each([
    ['demo', 'B1:B2313', 3],
    ['demo', 'B1:B2313', 50],
    ['lazy', 'F3:F2312', 3],
    ['lazy', 'B1:B2313', 50],
  ] as const)('topValues=%s %s %s', async (mode, range, topValues) => {
    await expectAggregate(mode as 'demo' | 'lazy', range, true, topValues)
  })

  it('topValues=0 → 无 top 行', async () => {
    const { deps } = demoDeps()
    const result = await runTool(deps, 'aggregate_range', { range: 'B1:B2313', topValues: 0 })
    expect(result.isError).toBeFalsy()
    expect(result.output).not.toContain('top values by frequency')
  })

  it('topValues 负数钳制为 0', async () => {
    const { deps } = demoDeps()
    const result = await runTool(deps, 'aggregate_range', { range: 'B1:B2313', topValues: -5 })
    expect(result.isError).toBeFalsy()
    expect(result.output).not.toContain('top values by frequency')
  })
})

// ---------------------------------------------------------------------------
// 拒绝路径
// ---------------------------------------------------------------------------

describe('aggregate 拒绝路径', () => {
  it.each([
    ['', 'range must be a non-empty string'],
    ['ABC', 'names whole columns or rows'],
    ['1:5', 'names whole columns or rows'],
    ['A1:B2:C3', 'range must be cell addresses'],
    ['A0', 'range must be cell addresses'],
  ] as const)('非法输入 %s 被拒', async (raw: string, fragment: string) => {
    await expectAggregateErr('demo', raw, fragment)
  })

  it('lazy: 完全越界被拒', async () => {
    await expectAggregateErr('lazy', 'A3000:G3000', 'outside the sheet data extent')
  })

  it('lazy: structural ops 后拒绝', async () => {
    await expectAggregateErr(
      'lazy',
      'B1:B10',
      'structural changes',
      { structuralOps: [{ kind: 'insert-rows', index: 100, count: 1 }] },
    )
  })

  it('lazy: 索引未完成被拒', async () => {
    await expectAggregateErr('lazy', 'B1:B2313', 'still being indexed', {
      sidecar: { indexedThroughRow: 100 },
    })
  })

  it('lazy: 部分索引——索引内的段可以聚合', async () => {
    const { deps } = lazyDeps({ sidecar: { indexedThroughRow: 2000 } })
    const result = await runTool(deps, 'aggregate_range', { range: 'B1:B200' })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('Statistics for B1:B200:')
  })

  it('demo: unknown sheetId 被拒', async () => {
    const { deps } = demoDeps()
    const result = await runTool(deps, 'aggregate_range', { range: 'B1:B10', sheetId: 'nope' })
    expect(result.isError).toBe(true)
    expect(result.output).toContain('Unknown sheet: nope')
  })

  it('demo: 显式 sheetId 通过', async () => {
    const { deps } = demoDeps()
    const result = await runTool(deps, 'aggregate_range', { range: 'B1:B10', sheetId: SHEET_ID })
    expect(result.isError).toBeFalsy()
  })
})

// ---------------------------------------------------------------------------
// 引擎无缓存
// ---------------------------------------------------------------------------

describe('aggregate engineHasCache=false', () => {
  it('demo F 列退化为全空（non-empty 0）', async () => {
    const { deps } = demoDeps({ engineHasCache: false })
    const result = await runTool(deps, 'aggregate_range', { range: 'F3:F2312' })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('cells: 2,310, non-empty: 0')
    expect(result.output).not.toContain('numeric cells')
  })

  it('lazy 不受引擎缓存影响（sidecar 直读文件缓存值）', async () => {
    const withCache = await runTool(lazyDeps().deps, 'aggregate_range', { range: 'F3:F2312' })
    const without = await runTool(
      lazyDeps({ engineHasCache: false }).deps,
      'aggregate_range',
      { range: 'F3:F2312' },
    )
    expect(without.output).toBe(withCache.output)
  })
})
