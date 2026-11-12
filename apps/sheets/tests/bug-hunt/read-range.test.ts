/**
 * read_range 边界猎捕（~85 题）：
 * - 数据区逐格严格断言（oracle 从 fixture 独立复算，含 demo/lazy 公式格渲染差异）
 * - 2000 格上限边界（2000 过 / 2001 拒）
 * - lazy extent 检查（行/列越界、边界格）
 * - 绝对引用 $A$1、反序范围、空白容错、非法输入
 * - 孤儿公式行 1888（F 列单独有 =EVALUATE(E1888)）
 * - engineHasCache=false（引擎无缓存 → 公式格渲染纯公式文本）
 */
import { afterEach, describe, expect, it } from 'vitest'

import {
  columnLabel,
  formatAddress,
  parseRange,
} from '@genoffice/xlsx-gateway/domain/cell-address'

import {
  COLS,
  demoDeps,
  expectedProfileParts,
  expectedReadRows,
  LAST_COL,
  lazyDeps,
  restoreGlobals,
  ROWS,
  runTool,
  type DepsBundle,
} from './harness'

afterEach(() => {
  restoreGlobals()
})

function makeDeps(mode: 'demo' | 'lazy', engineHasCache = true): DepsBundle {
  return mode === 'demo' ? demoDeps({ engineHasCache }) : lazyDeps({ engineHasCache })
}

function metaLine(_mode: 'demo' | 'lazy', normalized: string): string {
  // F29: demo 模式现在也报权威 extent（harness 对齐真 app demo 分支
  // workbook-readers.ts L399-412），与 lazy 同串——原 demo 分支的
  // "extent is unknown" 是替身丢 rows/columns 的旧行为（R13 幻觉 2900 行根因）。
  return `Read metadata: requested range ${normalized}; authoritative worksheet data extent A1:${LAST_COL}${ROWS} (${ROWS} worksheet rows including any header). Do not infer total rows or records from the requested range.`
}

/** 有效范围：数据区 + profile 计数全量断言；返回完整输出供对照 */
async function expectReadOk(
  mode: 'demo' | 'lazy',
  raw: string,
  engineHasCache = true,
): Promise<string> {
  const { deps } = makeDeps(mode, engineHasCache)
  const bounds = parseRange(raw.trim().toUpperCase())
  const result = await runTool(deps, 'read_range', { range: raw })
  expect(result.isError).toBeFalsy()
  const normalized = `${formatAddress(bounds.startRow, bounds.startColumn)}:${formatAddress(bounds.endRow, bounds.endColumn)}`
  const header = [
    '',
    ...Array.from(
      { length: bounds.endColumn - bounds.startColumn + 1 },
      (_, i) => columnLabel(bounds.startColumn + i),
    ),
  ].join('\t')
  const rows = expectedReadRows(
    bounds.startRow,
    bounds.endRow,
    bounds.startColumn,
    bounds.endColumn,
    mode,
    engineHasCache,
  )
  // 输出结构：meta \n header \n 数据行... \n（空行）\n Column profile...
  // （tools.ts rows.push('', columnProfileLine(...))）
  const prefix = [metaLine(mode, normalized), header, ...rows].join('\n')
  expect(result.output.startsWith(prefix)).toBe(true)
  const rest = result.output.slice(prefix.length)
  expect(rest.startsWith('\n\nColumn profile (this range): ')).toBe(true)
  for (const part of expectedProfileParts(
    bounds.startRow,
    bounds.endRow,
    bounds.startColumn,
    bounds.endColumn,
    mode,
    engineHasCache,
  )) {
    expect(rest).toContain(part)
  }
  return result.output
}

async function expectReadErr(mode: 'demo' | 'lazy', raw: string, fragment: string): Promise<void> {
  const { deps } = makeDeps(mode)
  const result = await runTool(deps, 'read_range', { range: raw })
  expect(result.isError).toBe(true)
  expect(result.output).toContain(fragment)
}

// ---------------------------------------------------------------------------
// demo 模式（无 extent 检查）
// ---------------------------------------------------------------------------

describe('read_range demo: 单格', () => {
  it.each(['A1', 'B1', 'B33', 'E3', 'F3', 'F1888', 'A2313', 'G2313'])('单格 %s 逐格匹配', async (range: string) => {
    await expectReadOk('demo', range)
  })

  it('demo: fixture 外单格（H1/Z5/A3000）被 extent 边界拒绝（F29：demo 现在有 extent）', async () => {
    // 旧行为：demo 无 extent → 越界读渲染为空行。F29 后 demo 与 lazy 同款
    // bounds-check，越界读被拒（模型不该读数据区外的格子，拒绝是正确信号）。
    for (const range of ['H1', 'Z5', 'A3000']) {
      await expectReadErr('demo', range, `outside the worksheet data extent A1:${LAST_COL}${ROWS}`)
    }
  })

  it('demo: F3 公式格无引擎值 → 渲染纯公式文本', async () => {
    const { deps } = demoDeps()
    const result = await runTool(deps, 'read_range', { range: 'F3' })
    const line = result.output.split('\n').find((l) => l.startsWith('3\t'))!
    expect(line).toBe(`3\t=EVALUATE(E3)`)
  })
})

describe('read_range demo: 行/列/矩形', () => {
  it.each([
    'A5:G5',
    'A1888:G1888', // 孤儿公式行整行
    'B1:B50',
    'B2260:B2313', // 表尾段
    'F1:F50',
    'F2263:F2312',
    'E1:E20',
    'C1:C100',
    'D1:D100',
    'G1:G50',
    'A1:G20',
    'B2:E100',
    'D100:F200',
    'B1:G50',
    // AA1:AA5（fixture 外列）F29 起被 extent 边界拒绝——越界读断言移到
    // 上面"fixture 外单格"用例与 lazy 越界被拒列表
  ] as const)('范围 %s 逐格匹配', async (range: string) => {
    await expectReadOk('demo', range)
  })
})

describe('read_range demo: 输入语法', () => {
  it.each([['$A$1:$C$5'], ['$B$2:$B$10'], ['G3:A1'], [' A1 : G3 ']] as const)('语法变体 %s 归一后逐格匹配', async (raw: string) => {
    await expectReadOk('demo', raw)
  })

  it.each([
    ['', 'range must be a non-empty string'],
    ['   ', 'range must be a non-empty string'],
    ['ABC', 'names whole columns or rows'],
    ['A1:B2:C3', 'range must be cell addresses'],
    ['1:5', 'names whole columns or rows'],
    ['A0', 'range must be cell addresses'],
    ['A', 'names whole columns or rows'],
    [':A1', 'range must be cell addresses'],
    ['A1:', 'range must be cell addresses'],
  ] as const)('非法输入 %s 被拒', async (raw: string, fragment: string) => {
    await expectReadErr('demo', raw, fragment)
  })
})

describe('read_range demo: 2000 格上限', () => {
  it('恰好 2000 格（A1:A2000）通过', async () => {
    await expectReadOk('demo', 'A1:A2000')
  })

  it.each([['A1:A2001'], ['G1:G2313'], ['A1:G3000']] as const)('超限 %s 被拒', async (range: string) => {
    await expectReadErr('demo', range, 'more than 2000 cells')
  })
})

// ---------------------------------------------------------------------------
// lazy 模式（extent A1:G2313）
// ---------------------------------------------------------------------------

describe('read_range lazy: 单格', () => {
  it.each(['A1', 'G2313', 'A2313', 'E3'])('单格 %s 逐格匹配（公式格带缓存值）', async (range: string) => {
    await expectReadOk('lazy', range)
  })

  it('lazy: F3 公式格渲染 缓存值 (公式)', async () => {
    const { deps } = lazyDeps()
    const result = await runTool(deps, 'read_range', { range: 'F3' })
    const line = result.output.split('\n').find((l) => l.startsWith('3\t'))!
    expect(line).toContain('(=EVALUATE(E3))')
    expect(line).not.toMatch(/^3\t=EVALUATE\(E3\)\t*$/) // 不是纯公式文本
  })

  it('lazy: 孤儿行 F1888 = 缓存 #VALUE! + 公式注记', async () => {
    const { deps } = lazyDeps()
    const result = await runTool(deps, 'read_range', { range: 'F1888' })
    const line = result.output.split('\n').find((l) => l.startsWith('1888\t'))!
    expect(line).toContain('#VALUE!')
    expect(line).toContain('(=EVALUATE(E1888))')
  })
})

describe('read_range lazy: 行/列/矩形', () => {
  it.each([
    'A5:G5',
    'A1888:G1888',
    'B1:B50',
    'B2300:B2313',
    'F1:F50',
    'F2263:F2312',
    'G2300:G2313',
    'A1:G20',
    'B2:E100',
    'D100:F200',
    'A2300:G2313',
  ] as const)('范围 %s 逐格匹配', async (range) => {
    await expectReadOk('lazy', range)
  })
})

describe('read_range lazy: extent 检查', () => {
  // 全部选 ≤2000 格的范围，避免先撞 2000 格上限（上限检查先于 extent 检查）
  it.each([
    ['A2314:B2314', `outside the worksheet data extent A1:${LAST_COL}${ROWS}`],
    ['A2314:G2314', `outside the worksheet data extent A1:${LAST_COL}${ROWS}`],
    ['H1:H5', `outside the worksheet data extent A1:${LAST_COL}${ROWS}`],
    ['AA1:AA5', `outside the worksheet data extent A1:${LAST_COL}${ROWS}`],
    ['B3000:C3000', `outside the worksheet data extent A1:${LAST_COL}${ROWS}`],
  ] as const)('越界 %s 被拒', async (range: string, fragment: string) => {
    await expectReadErr('lazy', range, fragment)
  })

  it('边界格 G2313 / A2313 单格可读（extent 闭区间）', async () => {
    await expectReadOk('lazy', 'G2313')
    await expectReadOk('lazy', 'A2313')
  })

  it.each([['A1:G2313'], ['A1:A2001']] as const)('2000 上限优先于 extent：超限 %s', async (range: string) => {
    await expectReadErr('lazy', range, 'more than 2000 cells')
  })
})

describe('read_range lazy: 输入语法', () => {
  it.each([['$A$1:$C$5'], ['G3:A1']] as const)('语法变体 %s 归一后逐格匹配', async (raw: string) => {
    await expectReadOk('lazy', raw)
  })

  it.each([
    ['', 'range must be a non-empty string'],
    ['ABC', 'names whole columns or rows'],
    ['A0', 'range must be cell addresses'],
    [':A1', 'range must be cell addresses'],
  ] as const)('非法输入 %s 被拒', async (raw: string, fragment: string) => {
    await expectReadErr('lazy', raw, fragment)
  })
})

// ---------------------------------------------------------------------------
// 引擎无缓存变体（公式格显示退化为纯公式文本）
// ---------------------------------------------------------------------------

describe('read_range engineHasCache=false', () => {
  it.each([['demo'], ['lazy']] as const)('F3 退化为纯公式文本（%s）', async (mode: 'demo' | 'lazy') => {
    const { deps } = makeDeps(mode, false)
    const result = await runTool(deps, 'read_range', { range: 'F3' })
    const line = result.output.split('\n').find((l) => l.startsWith('3\t'))!
    expect(line).toBe(`3\t=EVALUATE(E3)`)
  })

  it.each([['demo'], ['lazy']] as const)('孤儿行 1888 退化为纯公式文本（%s）', async (mode: 'demo' | 'lazy') => {
    const { deps } = makeDeps(mode, false)
    const result = await runTool(deps, 'read_range', { range: 'A1888:G1888' })
    const line = result.output.split('\n').find((l) => l.startsWith('1888\t'))!
    const cells = line.split('\t')
    for (const idx of [1, 2, 3, 4, 5, 7]) expect(cells[idx]).toBe('')
    expect(cells[6]).toBe('=EVALUATE(E1888)')
  })

  it('lazy 无缓存：B 列（普通格）不受影响', async () => {
    const withCache = await expectReadOk('lazy', 'B1:B10')
    const { deps } = lazyDeps({ engineHasCache: false })
    const result = await runTool(deps, 'read_range', { range: 'B1:B10' })
    expect(result.isError).toBeFalsy()
    expect(result.output).toBe(withCache)
  })
})
