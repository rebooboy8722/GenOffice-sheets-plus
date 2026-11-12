/**
 * propose_operations 写侧猎捕（~90 题）：
 * - handler 输入校验（operations/summary 形状）
 * - zod schema 拒绝（describeOperationErrors：missing 字段 / unknown field hint / 小写地址 / 空 format）
 * - 批级校验（refine 混批 / set_range 四组 / 全批 2000 格上限 / fill 几何 / 200,000 上限）
 * - demo 路径（不支持的 op / convert 三态 / expectedValue CAS / 重名 / 最后一个 sheet）
 * - lazy 路径（structural 四段越界 / 未加载 gate / unknown sheet 后缀 / fill 源未加载）
 * - apply 阶段失败文案（UNCHANGED / MID-BATCH / undoDropped / unknown reason）
 * - adapter 直测（tx 重复 / CAS / undo）
 * - 写闭环：真实 InMemoryWorkbookAdapter 读回断言 + 输出逐字断言 +
 *   回读细节（10 格上限 / 跨 sheet 前缀 / ⚠️ 错误值行 / fill/copy 角格重复行）
 *
 * 已知 stub 边界：demo 引擎不计算公式（新公式回读为 still computing，⚠️ 行靠
 * setEngineOverlay 模拟引擎重算值）；lazy apply 走 {ok:true} stub（只验校验链
 * 与回读渲染，不验 lazy 写闭环）。
 */
import { afterEach, describe, expect, it } from 'vitest'

import { InMemoryWorkbookAdapter } from '@genoffice/xlsx-gateway/domain/in-memory-workbook'
import { workbookCommandBatchSchema } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

import { describeBatchValidationError } from '../../src/renderer/plan-operations'

import {
  demoDeps,
  escapeCellText,
  lazyDeps,
  restoreGlobals,
  runTool,
  SHEET_ID,
  SHEET_NAME,
  type DemoDepsOptions,
} from './harness'

type CellState = { value: string | number | boolean | null; formula?: string | undefined }

afterEach(() => {
  restoreGlobals()
})

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------

function live(options: DemoDepsOptions = {}): {
  deps: ReturnType<typeof demoDeps>['deps']
  adapter: InMemoryWorkbookAdapter
  setEngineOverlay: (v: Record<string, string | number | null>) => void
  initial: Record<string, CellState>
} {
  const bundle = demoDeps({ live: true, ...options })
  if (!bundle.adapter || !bundle.setEngineOverlay) throw new Error('live bundle incomplete')
  return {
    deps: bundle.deps,
    adapter: bundle.adapter,
    setEngineOverlay: bundle.setEngineOverlay,
    initial: bundle.adapter.getSnapshot().sheets.find((s) => s.id === SHEET_ID)?.cells ?? {},
  }
}

function cellOf(bundle: { adapter: InMemoryWorkbookAdapter }, address: string, sheetId = SHEET_ID) {
  return bundle.adapter.getSnapshot().sheets.find((s) => s.id === sheetId)?.cells[address]
}

/** 复刻 tools.ts formatCellScalar（summary 里 before/after 的渲染） */
function scalar(cell: CellState | undefined): string {
  const c: CellState = cell ?? { value: null }
  if (c.formula) {
    return c.value === null || c.value === undefined
      ? escapeCellText(c.formula)
      : `${escapeCellText(String(c.value))} (${escapeCellText(c.formula)})`
  }
  if (c.value === null) return '(empty)'
  return escapeCellText(String(c.value))
}

/** 单条 cellChanges 期望：`<addr>: <before> → <after>` */
function changeLine(initial: Record<string, CellState>, address: string, after: CellState): string {
  return `${address}: ${scalar(initial[address])} → ${scalar(after)}`
}

const UNDO_HINT = 'Auto-applied'

function expectAutoApplied(output: string, opCount: number, summary: string): void {
  expect(output).toBe(
    `Auto-applied ${opCount} change(s) (undo via the side panel [Undo] button or ⌘Z): ${summary}`,
  )
}

/** Rejected 包装（proposeOperations 层错误）逐字断言 */
async function expectRejected(
  bundle: { deps: Parameters<typeof runTool>[0] },
  operations: unknown[],
  errorText: string,
  extra?: Record<string, unknown>,
): Promise<void> {
  const result = await runTool(bundle.deps, 'propose_operations', {
    operations,
    summary: 't',
    ...extra,
  })
  expect(result.isError).toBe(true)
  expect(result.output).toBe(
    `Rejected — none of the ${operations.length} operation(s) were applied (a batch is all-or-nothing): ${errorText}`,
  )
}

// ---------------------------------------------------------------------------
// 1. handler 输入校验
// ---------------------------------------------------------------------------

describe('propose_operations handler 输入校验', () => {
  it.each([
    ['operations 是字符串', { operations: 'set_cell', summary: 't' }],
    ['operations 空数组', { operations: [], summary: 't' }],
    ['operations 缺失', { summary: 't' }],
  ])('%s 被拒', async (_name, input) => {
    const { deps } = live()
    const result = await runTool(deps, 'propose_operations', input as Record<string, unknown>)
    expect(result.isError).toBe(true)
    expect(result.output).toBe('operations must be a non-empty array')
  })

  it.each([
    ['summary 空串', { operations: [{ op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 1 }], summary: '   ' }],
    ['summary 缺失', { operations: [{ op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 1 }] }],
    ['summary 是数字', { operations: [{ op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 1 }], summary: 42 }],
  ])('%s 被拒', async (_name, input) => {
    const { deps } = live()
    const result = await runTool(deps, 'propose_operations', input as Record<string, unknown>)
    expect(result.isError).toBe(true)
    expect(result.output).toBe('summary must not be empty')
  })
})

// ---------------------------------------------------------------------------
// 2. zod schema 拒绝（describeOperationErrors）
// ---------------------------------------------------------------------------

const ZOD_PREFIX =
  'Rejected — none of the %N% operation(s) were applied (a batch is all-or-nothing). Fix the operations below and resubmit the whole batch, including the ones that were valid:\n'

async function expectZodRejection(
  operations: unknown[],
  ...lineFragments: string[]
): Promise<void> {
  const { deps } = live()
  const result = await runTool(deps, 'propose_operations', { operations, summary: 't' })
  expect(result.isError).toBe(true)
  expect(result.output.startsWith(ZOD_PREFIX.replace('%N%', String(operations.length))))
    .toBe(true)
  for (const fragment of lineFragments) {
    expect(result.output).toContain(fragment)
  }
}

describe('propose_operations zod 拒绝', () => {
  it('set_cell 缺 value（union 类型）→ missing value（bug-hunt P1 修复后）', async () => {
    // 修复前 union 缺失走 invalid_union 分支，输出 "value: Invalid input"
    await expectZodRejection(
      [{ op: 'set_cell', sheetId: SHEET_ID, address: 'G5' }],
      '- operations[0] (set_cell): missing value',
    )
  })

  it('set_cell 缺 address → missing address (expected string)', async () => {
    await expectZodRejection(
      [{ op: 'set_cell', sheetId: SHEET_ID, value: 1 }],
      '- operations[0] (set_cell): missing address (expected string)',
    )
  })

  it('unknown field 提示真实字段表（与 missing value 同列输出）', async () => {
    await expectZodRejection(
      [{ op: 'set_cell', sheetId: SHEET_ID, address: 'G5', col: 'A' }],
      '- operations[0] (set_cell): missing value; unknown field(s) col — set_cell takes: sheetId, address, value, expectedValue',
    )
  })

  it('set_formula 公式不以 = 开头', async () => {
    await expectZodRejection(
      [{ op: 'set_formula', sheetId: SHEET_ID, address: 'G5', formula: '1+1' }],
      '- operations[0] (set_formula): formula: ',
    )
  })

  it('未知 op 名', async () => {
    await expectZodRejection([{ op: 'bogus' }], '- operations[0] (bogus): op: ')
  })

  it('insert_rows count 超 500 上限', async () => {
    await expectZodRejection(
      [{ op: 'insert_rows', sheetId: SHEET_ID, row: 100, count: 501 }],
      '- operations[0] (insert_rows): count: ',
    )
  })

  it('format_range 空 format → refine 文案', async () => {
    await expectZodRejection(
      [{ op: 'format_range', sheetId: SHEET_ID, range: 'A1:B2', format: {} }],
      '- operations[0] (format_range): format: format must set at least one property',
    )
  })

  it('第二个 op 非法 → 只列 operations[1]，且强调重提整批', async () => {
    await expectZodRejection(
      [
        { op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 1 },
        { op: 'set_formula', sheetId: SHEET_ID, address: 'G6', formula: '1+1' },
      ],
      'including the ones that were valid',
      '- operations[1] (set_formula): formula: ',
    )
    // operations[0] 合法，不应出现在错误列表
    const { deps } = live()
    const result = await runTool(deps, 'propose_operations', {
      operations: [
        { op: 'set_cell', sheetId: SHEET_ID, address: 'G7', value: 1 },
        { op: 'set_formula', sheetId: SHEET_ID, address: 'G6', formula: '1+1' },
      ],
      summary: 't',
    })
    expect(result.output).not.toContain('operations[0]')
  })

  it('小写地址被 schema 拒绝（regex 全大写）', async () => {
    await expectZodRejection(
      [{ op: 'set_cell', sheetId: SHEET_ID, address: 'g5', value: 1 }],
      '- operations[0] (set_cell): address: ',
    )
  })
})

// ---------------------------------------------------------------------------
// 3. 批级校验拒绝（demo live，Rejected 包装逐字）
// ---------------------------------------------------------------------------

describe('propose_operations 批级校验（demo）', () => {
  it('unknown sheetId', async () => {
    await expectRejected(
      live(),
      [{ op: 'set_cell', sheetId: 'nope', address: 'G5', value: 1 }],
      'Unknown sheet: nope',
    )
  })

  it('structural × cell 混批被 refine 拒绝', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 1 },
        { op: 'insert_rows', sheetId: SHEET_ID, row: 100, count: 1 },
      ],
      summary: 't',
    })
    expect(result.isError).toBe(true)
    expect(result.output).toContain(
      'Structural operations (rows/columns/sheets) and cell edits must be proposed in separate batches.',
    )
    // O1 修复：批级 zod 拒绝不再吐 JSON.stringify(issues) blob（模型读不懂→白烧轮次），
    // 而是渲染 issue.message 本身（人话）
    expect(result.output).not.toContain('"code": "custom"')
    expect(result.output).not.toContain('[{')
  })

  it('describeBatchValidationError：ZodError → 人话句；非 ZodError → null（O1 修复单元）', () => {
    const parsed = workbookCommandBatchSchema.safeParse({
      dslVersion: 1,
      transactionId: 't1',
      baseRevision: 0,
      summary: 's',
      // 字段齐全 → 唯一 issue 来自 refine（混批），而非字段校验噪音
      operations: [
        { op: 'set_cell', sheetId: 's', address: 'A1', value: 1 },
        { op: 'insert_rows', sheetId: 's', row: 1, count: 1 },
      ],
    })
    expect(parsed.success).toBe(false)
    const text = describeBatchValidationError(parsed.error)
    expect(text).toBe(
      'Structural operations (rows/columns/sheets) and cell edits must be proposed in separate batches.',
    )
    // 对照组：Error 走 fallback 分支（error.message），不经过这个 helper
    expect(describeBatchValidationError(new Error('boom'))).toBeNull()
    expect(describeBatchValidationError('plain string')).toBeNull()
  })

  it('rename_sheet × set_cell 混批放行（刻意设计：rename 不占单元格坐标）', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 1 },
        { op: 'rename_sheet', sheetId: SHEET_ID, name: '工程量' },
      ],
      summary: 'mix',
    })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain(UNDO_HINT)
    expect(cellOf(bundle, 'G5')?.value).toBe(1)
    expect(bundle.adapter.getSnapshot().sheets[0]?.name).toBe('工程量')
  })

  it('set_range 锯齿 values', async () => {
    await expectRejected(
      live(),
      [{ op: 'set_range', sheetId: SHEET_ID, start: 'G5', values: [[1, 2], [3]] }],
      'set_range values must be rectangular: values[0] has 2 cell(s) but values[1] has 1. Use null for cells that should be cleared, or split into separate set_range operations.',
    )
  })

  it('set_range 缺 start/range', async () => {
    await expectRejected(
      live(),
      [{ op: 'set_range', sheetId: SHEET_ID, values: [[1]] }],
      'set_range needs "start" — the top-left target cell, like "B2".',
    )
  })

  it('set_range start 与 range 不一致', async () => {
    await expectRejected(
      live(),
      [{ op: 'set_range', sheetId: SHEET_ID, start: 'G5', range: 'G6', values: [[1]] }],
      'set_range received both start G5 and range G6, which disagree on the top-left cell — pass only one of them.',
    )
  })

  it('set_range range 尺寸与 values 不符', async () => {
    await expectRejected(
      live(),
      [{ op: 'set_range', sheetId: SHEET_ID, range: 'G5:G6', values: [[1], [2], [3]] }],
      'set_range range G5:G6 spans 2×1 cells but values is 3 row(s) × 1 cell(s) — make them match, or give "start" (the top-left cell) instead.',
    )
  })

  it('全批展开超 2000 格上限（500×5 = 2500）', async () => {
    await expectRejected(
      live(),
      [{ op: 'set_range', sheetId: SHEET_ID, start: 'G1', values: Array.from({ length: 500 }, () => [1, 2, 3, 4, 5]) }],
      'The batch expands to more than 2000 cell edits.',
    )
  })

  it('fill_range 源超 2000 格', async () => {
    await expectRejected(
      live(),
      [{ op: 'fill_range', sheetId: SHEET_ID, source: 'A1:G286', target: 'A1:A2002' }],
      'fill_range source covers more than 2000 cells.',
    )
  })

  it('fill_range 目标非源整数倍', async () => {
    await expectRejected(
      live(),
      [{ op: 'fill_range', sheetId: SHEET_ID, source: 'A1:B2', target: 'A1:C3' }],
      'fill_range target A1:C3 (3×3) is not a whole multiple of source A1:B2 (2×2) — adjust the target so the source tiles it exactly.',
    )
  })

  it('fill_range 源与目标重叠（非同角）', async () => {
    await expectRejected(
      live(),
      [{ op: 'fill_range', sheetId: SHEET_ID, source: 'A1:B2', target: 'B2:C3' }],
      'fill_range source and target overlap — start the target at the source cell (fill-down/right includes the source as the first tile) or keep them disjoint.',
    )
  })

  it('fill_range demo 目标超 2000 格（adapter 上限）', async () => {
    await expectRejected(
      live(),
      [{ op: 'fill_range', sheetId: SHEET_ID, source: 'A1', target: 'A1:A2001' }],
      'fill_range on an in-memory workbook is limited to 2000 cells per operation (imported xlsx files allow up to 200,000).',
    )
  })

  it.each([
    ['find_replace', 'find_replace covers more than 200,000 cells — replace in several ranges.',
      { op: 'find_replace', sheetId: SHEET_ID, range: 'A1:G40000', find: 'x', replace: 'y' }],
    ['sort_range', 'sort_range covers more than 200,000 cells — sort a smaller range.',
      { op: 'sort_range', sheetId: SHEET_ID, range: 'A1:G40000', byColumn: 'A', order: 'asc' }],
  ])('%s 超 200,000 格上限', async (_name, errorText, operation) => {
    await expectRejected(live(), [operation], errorText)
  })

  it('clear_range 超 200,000 格上限（超 F11 守卫阈值 → 先确认才轮到 200,000 校验）', async () => {
    // A1:G40000 = 280,000 cells ≥ 5,000 → destructiveOperationNotes 先拦；
    // 这个用例测的是 200,000 上限本身，所以带 confirmDestructive 放行守卫。
    await expectRejected(
      live(),
      [{ op: 'clear_range', sheetId: SHEET_ID, range: 'A1:G40000' }],
      'clear_range covers more than 200,000 cells — clear it in several ranges.',
      { confirmDestructive: true },
    )
  })

  it('find_replace 超 2000 格走 range-level → demo adapter 上限拒绝', async () => {
    await expectRejected(
      live(),
      [{ op: 'find_replace', sheetId: SHEET_ID, range: 'A1:G2313', find: 'x', replace: 'y' }],
      'find_replace on an in-memory workbook is limited to 2000 cells per operation (imported xlsx files allow up to 200,000).',
    )
  })

  it('sort_range 超 2000 格走 range-level → demo adapter 上限拒绝', async () => {
    await expectRejected(
      live(),
      [{ op: 'sort_range', sheetId: SHEET_ID, range: 'A1:G2313', byColumn: 'A', order: 'asc' }],
      'sort_range on an in-memory workbook is limited to 2000 cells per operation (imported xlsx files allow up to 200,000).',
    )
  })
})

// ---------------------------------------------------------------------------
// 4. demo 路径专属拒绝
// ---------------------------------------------------------------------------

describe('propose_operations demo 专属拒绝', () => {
  it('set_cell expectedValue 不匹配', async () => {
    await expectRejected(
      live(),
      [{ op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 1, expectedValue: 2 }],
      `${SHEET_NAME}!G5 no longer has the expected value.`,
    )
  })

  it('set_formula expectedFormula 不匹配', async () => {
    await expectRejected(
      live(),
      [{ op: 'set_formula', sheetId: SHEET_ID, address: 'F100', formula: '=1', expectedFormula: '=X' }],
      `${SHEET_NAME}!F100 no longer has the expected formula.`,
    )
  })

  it('rename_sheet 改成已存在名（含自己）被拒', async () => {
    await expectRejected(
      live(),
      [{ op: 'rename_sheet', sheetId: SHEET_ID, name: SHEET_NAME }],
      `A sheet named "${SHEET_NAME}" already exists.`,
    )
  })

  it('add_sheet 重名被拒', async () => {
    await expectRejected(
      live(),
      [{ op: 'add_sheet', name: SHEET_NAME }],
      `A sheet named "${SHEET_NAME}" already exists.`,
    )
  })

  it('delete_sheet 最后一个 sheet 被拒（F11：delete_sheet 无条件破坏性 → 确认后仍被"需保留一个 sheet"拒）', async () => {
    await expectRejected(
      live(),
      [{ op: 'delete_sheet', sheetId: SHEET_ID }],
      'The workbook needs at least one sheet.',
      { confirmDestructive: true },
    )
  })

  it.each([
    ['duplicate_sheet', { op: 'duplicate_sheet', sheetId: SHEET_ID }],
    ['set_sheet_hidden', { op: 'set_sheet_hidden', sheetId: SHEET_ID, hidden: true }],
    ['move_sheet', { op: 'move_sheet', sheetId: SHEET_ID, position: 1 }],
  ] as const)('%s 在 demo 工作簿不支持', async (op, operation) => {
    await expectRejected(
      live(),
      [operation],
      `${op} needs an imported xlsx file — the demo workbook does not support it.`,
    )
  })

  it('convert_to_values 无公式格', async () => {
    await expectRejected(
      live(),
      [{ op: 'convert_to_values', sheetId: SHEET_ID, range: 'A1:A5' }],
      'The convert_to_values range contains no formula cells — nothing to convert.',
    )
  })

  it('convert_to_values 超 2000 格', async () => {
    await expectRejected(
      live(),
      [{ op: 'convert_to_values', sheetId: SHEET_ID, range: 'F1:F2001' }],
      'convert_to_values on an in-memory workbook is limited to 2000 cells per operation (imported xlsx files allow up to 200,000).',
    )
  })

  it('convert_to_values 与同范围 set_formula 混批被拒', async () => {
    await expectRejected(
      live(),
      [
        { op: 'convert_to_values', sheetId: SHEET_ID, range: 'F3:F4' },
        { op: 'set_formula', sheetId: SHEET_ID, address: 'F3', formula: '=1' },
      ],
      'convert_to_values on F3:F4 cannot share a batch with a set_formula that writes formulas into that range — the writes would land after (or invisibly to) the convert. Propose the formula writes first, verify the results, then convert in a separate batch.',
    )
  })

  it('copy_range filter 无匹配行', async () => {
    await expectRejected(
      live(),
      [
        {
          op: 'copy_range',
          sheetId: SHEET_ID,
          source: 'B2:B5',
          target: 'G200',
          filterColumn: 'B',
          filterValues: ['不存在'],
        },
      ],
      'copy_range filter matched no source rows — none of the B cells equal 不存在 (matching is trimmed and case-insensitive).',
    )
  })
})

// ---------------------------------------------------------------------------
// 5. lazy 路径专属拒绝
// ---------------------------------------------------------------------------

describe('propose_operations lazy 拒绝', () => {
  async function lazyRejected(
    operations: unknown[],
    errorText: string,
    options?: Parameters<typeof lazyDeps>[0],
    extra?: Record<string, unknown>,
  ): Promise<void> {
    const { deps } = lazyDeps(options ?? {})
    const result = await runTool(deps, 'propose_operations', {
      operations,
      summary: 't',
      ...extra,
    })
    expect(result.isError).toBe(true)
    expect(result.output).toBe(
      `Rejected — none of the ${operations.length} operation(s) were applied (a batch is all-or-nothing): ${errorText}`,
    )
  }

  it('unknown sheetId 带 get_workbook_context 提示', async () => {
    await lazyRejected(
      [{ op: 'set_cell', sheetId: 'nope', address: 'G5', value: 1 }],
      'Unknown sheet: nope (use an id from get_workbook_context)',
    )
  })

  it('insert_rows 越界', async () => {
    await lazyRejected(
      [{ op: 'insert_rows', sheetId: SHEET_ID, row: 2314, count: 1 }],
      'insert_rows: row=2314 is beyond the sheet grid (2313 rows) — rows can only be inserted before an existing row (1–2313). To add data below the last row, write to the address directly — the grid grows automatically.',
    )
  })

  it('delete_rows 越界（count=200 超 F11 阈值 → 确认后仍被越界校验拒）', async () => {
    await lazyRejected(
      [{ op: 'delete_rows', sheetId: SHEET_ID, row: 2300, count: 200 }],
      'delete_rows: rows 2300–2499 extend beyond the sheet grid (2313 rows).',
      undefined,
      { confirmDestructive: true },
    )
  })

  it('insert_cols 越界（G 之后）', async () => {
    await lazyRejected(
      [{ op: 'insert_cols', sheetId: SHEET_ID, column: 'H', count: 1 }],
      'insert_cols: column H is beyond the sheet grid (last column G) — columns can only be inserted before an existing column (A–G). To put data in a new column past G, write to it directly (set_cell / set_range / set_formula) — the grid grows automatically.',
    )
  })

  it('delete_cols 越界', async () => {
    await lazyRejected(
      [{ op: 'delete_cols', sheetId: SHEET_ID, column: 'G', count: 2 }],
      'delete_cols: columns G + count=2 extend beyond the sheet grid (last column G).',
    )
  })

  it.each(['find_replace', 'sort_range'] as const)('%s 未加载被 gate（%s）', async (op) => {
    const operation =
      op === 'find_replace'
        ? { op, sheetId: SHEET_ID, range: 'B1:B10', find: 'x', replace: 'y' }
        : { op, sheetId: SHEET_ID, range: 'B1:B10', byColumn: 'B', order: 'asc' }
    // gate 读 state.loadedRanges + state.formulaMode（plan-operations L84-100），
    // 与 ensureRangeLoaded 回调无关
    await lazyRejected(
      [operation],
      `The ${op} range is not fully loaded yet — read_range it first (that loads it), then retry.`,
      { formulaMode: false, loadedRanges: 'none' },
    )
  })

  it('find_replace 超 2000 格豁免未加载 gate（formulaMode 下直接放行）', async () => {
    const { deps } = lazyDeps()
    const result = await runTool(deps, 'propose_operations', {
      operations: [{ op: 'find_replace', sheetId: SHEET_ID, range: 'A1:G2313', find: '疑问', replace: '疑点' }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain(UNDO_HINT)
  })

  it('fill_range 源未加载被拒（lazyRangeEditable：preloadComplete=false 且 loadedRanges 空）', async () => {
    await lazyRejected(
      [{ op: 'fill_range', sheetId: SHEET_ID, source: 'B1:B5', target: 'G1:G5' }],
      'The fill source is not loaded yet — read_range the source first (that loads it), then retry the fill.',
      { preloadComplete: false, loadedRanges: 'none' },
    )
  })

  it('copy_range 源越出 grid', async () => {
    await lazyRejected(
      [{ op: 'copy_range', sheetId: SHEET_ID, source: 'A3000:G3000', target: 'A1' }],
      'The source range A3000:G3000 extends beyond the source sheet grid (2313 rows × 7 columns) — copy only cells that exist.',
    )
  })
})

// ---------------------------------------------------------------------------
// 6. apply 阶段失败文案（failApply 覆盖三分支）
// ---------------------------------------------------------------------------

describe('propose_operations apply 失败文案', () => {
  it('UNCHANGED（非 partial，reason 不带尾句号）', async () => {
    const { deps } = live({ failApply: () => ({ ok: false, reason: 'boom' }) })
    const result = await runTool(deps, 'propose_operations', {
      operations: [{ op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 1 }],
      summary: 't',
    })
    expect(result.isError).toBe(true)
    expect(result.output).toBe(
      'Apply failed — the workbook is UNCHANGED: boom. ' +
        'Do not tell the user the changes were made; adjust the operations and retry, or explain the failure.',
    )
  })

  it('UNCHANGED reason 自带尾句号 → 归一为单句号（O2 修复）', async () => {
    const { deps } = live({
      failApply: () => ({ ok: false, reason: 'The workbook changed after preview.' }),
    })
    const result = await runTool(deps, 'propose_operations', {
      operations: [{ op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 1 }],
      summary: 't',
    })
    expect(result.isError).toBe(true)
    // O2 修复：reason 尾部的 '.' 先剥掉再接模板的 ". Do not tell…"，不再出 ".."
    expect(result.output).toBe(
      'Apply failed — the workbook is UNCHANGED: The workbook changed after preview. ' +
        'Do not tell the user the changes were made; adjust the operations and retry, or explain the failure.',
    )
    expect(result.output).not.toContain('..')
  })

  it('MID-BATCH（partial，整批一步 undo）', async () => {
    const { deps } = live({ failApply: () => ({ ok: false, reason: 'boom', partiallyApplied: true }) })
    const result = await runTool(deps, 'propose_operations', {
      operations: [{ op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 1 }],
      summary: 't',
    })
    expect(result.isError).toBe(true)
    expect(result.output).toBe(
      'Apply failed MID-BATCH — operations before the failing one were already committed: boom. ' +
        'Read the affected ranges to see the current state before continuing; ' +
        'the whole partial batch is one undo step (⌘Z / [Undo]).',
    )
  })

  it('MID-BATCH + undoDropped（⌘Z 不生效）', async () => {
    const { deps } = live({
      failApply: () => ({ ok: false, reason: 'boom', partiallyApplied: true, undoDropped: true }),
    })
    const result = await runTool(deps, 'propose_operations', {
      operations: [{ op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 1 }],
      summary: 't',
    })
    expect(result.isError).toBe(true)
    expect(result.output).toBe(
      'Apply failed MID-BATCH — operations before the failing one were already committed: boom. ' +
        'Read the affected ranges to see the current state before continuing; ' +
        'the committed part is too large for the undo history — ⌘Z will NOT revert it.',
    )
  })

  it('reason 缺失 → unknown reason', async () => {
    const { deps } = live({ failApply: () => ({ ok: false }) })
    const result = await runTool(deps, 'propose_operations', {
      operations: [{ op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 1 }],
      summary: 't',
    })
    expect(result.isError).toBe(true)
    expect(result.output).toContain('Apply failed — the workbook is UNCHANGED: unknown reason.')
  })
})

// ---------------------------------------------------------------------------
// 7. InMemoryWorkbookAdapter 直测
// ---------------------------------------------------------------------------

describe('InMemoryWorkbookAdapter 直测', () => {
  const tiny = () =>
    new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [{ id: 's1', name: 'S', cells: { A1: { value: 1 } } }],
    })

  it('同一 plan apply 两次 → CAS 先行拦截（观察项 O3：此场景 tx 重复检查不可达）', () => {
    const adapter = tiny()
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx1',
      baseRevision: 0,
      summary: 't',
      operations: [{ op: 'set_cell', sheetId: 's1', address: 'A1', value: 2 }],
    })
    adapter.apply(plan)
    // apply 先查 baseRevision（apply 已 revision+1 → 失配），
    // 后面的 committedTransactions 检查在"同一 plan 二次 apply"场景到不了
    expect(() => adapter.apply(plan)).toThrow('The workbook changed after preview.')
  })

  it('baseRevision 落后 → CAS 拒绝', () => {
    const adapter = tiny()
    adapter.apply(adapter.plan({
      dslVersion: 1,
      transactionId: 'tx1',
      baseRevision: 0,
      summary: 't',
      operations: [{ op: 'set_cell', sheetId: 's1', address: 'A1', value: 2 }],
    }))
    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx2',
        baseRevision: 0,
        summary: 't',
        operations: [{ op: 'set_cell', sheetId: 's1', address: 'A1', value: 3 }],
      }),
    ).toThrow('The workbook changed after the AI context was captured.')
  })

  it('undo 空栈被拒', () => {
    expect(() => tiny().undo()).toThrow('There is no transaction to undo.')
  })

  it('apply 后 undo 恢复原值', () => {
    const adapter = tiny()
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx1',
      baseRevision: 0,
      summary: 't',
      operations: [{ op: 'set_cell', sheetId: 's1', address: 'A1', value: 2 }],
    })
    adapter.apply(plan)
    expect(adapter.getSnapshot().sheets[0]?.cells['A1']?.value).toBe(2)
    adapter.undo()
    expect(adapter.getSnapshot().sheets[0]?.cells['A1']?.value).toBe(1)
  })
})

// ---------------------------------------------------------------------------
// 8. 写闭环：基础单格 op（真实 adapter 读回）
// ---------------------------------------------------------------------------

describe('propose_operations 写闭环（demo live 基础 op）', () => {
  it('set_cell 覆盖普通格：逐字 summary + 快照落盘', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'set_cell', sheetId: SHEET_ID, address: 'A100', value: '新内容' }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expectAutoApplied(
      result.output,
      1,
      changeLine(bundle.initial, 'A100', { value: '新内容' }),
    )
    expect(cellOf(bundle, 'A100')).toEqual({ value: '新内容' })
  })

  it('set_cell 写空格：before 渲染 (empty)', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 42 }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expectAutoApplied(result.output, 1, 'G5: (empty) → 42')
    expect(cellOf(bundle, 'G5')).toEqual({ value: 42 })
  })

  it('set_formula：summary 纯公式 → 回读 still computing（demo 引擎不算）', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'set_formula', sheetId: SHEET_ID, address: 'E100', formula: '=1+1' }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expect(result.output).toBe(
      `Auto-applied 1 change(s) (undo via the side panel [Undo] button or ⌘Z): ${changeLine(bundle.initial, 'E100', { value: null, formula: '=1+1' })}` +
        '\nFormula results: E100 = (still computing; verify with read_cells)',
    )
    expect(cellOf(bundle, 'E100')).toEqual({ value: null, formula: '=1+1' })
  })

  it('set_formula + 引擎回算错误值 → ⚠️ 行', async () => {
    const bundle = live()
    bundle.setEngineOverlay({ F100: '#DIV/0!' })
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'set_formula', sheetId: SHEET_ID, address: 'F100', formula: '=1/0' }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expect(result.output).toBe(
      `Auto-applied 1 change(s) (undo via the side panel [Undo] button or ⌘Z): ${changeLine(bundle.initial, 'F100', { value: null, formula: '=1/0' })}` +
        '\nFormula results: F100 = #DIV/0!' +
        '\n⚠️ Formula error values present — check references/divisors and fix them.',
    )
  })

  it('clear_cell 清公式格：summary 公式 → (empty)，无回读', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'clear_cell', sheetId: SHEET_ID, address: 'F100' }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expectAutoApplied(result.output, 1, `${changeLine(bundle.initial, 'F100', { value: null })}`)
    expect(result.output).not.toContain('Formula results')
    const cleared = cellOf(bundle, 'F100')
    expect(cleared?.formula).toBeUndefined()
    expect(cleared?.value ?? null).toBeNull()
  })

  it('set_range start+values（含 null 清空）：4 格 row-major 逐字 + 快照', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'set_range', sheetId: SHEET_ID, start: 'A100', values: [['x', null], [3, 'y']] },
      ],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    const summary = [
      changeLine(bundle.initial, 'A100', { value: 'x' }),
      changeLine(bundle.initial, 'B100', { value: null }),
      changeLine(bundle.initial, 'A101', { value: 3 }),
      changeLine(bundle.initial, 'B101', { value: 'y' }),
    ].join('; ')
    expectAutoApplied(result.output, 4, summary)
    expect(cellOf(bundle, 'A100')?.value).toBe('x')
    expect(cellOf(bundle, 'B100')?.value ?? null).toBeNull()
    expect(cellOf(bundle, 'A101')?.value).toBe(3)
    expect(cellOf(bundle, 'B101')?.value).toBe('y')
  })

  it('set_range "="-字符串展开成 set_formula：回读两格（2 行 × 1 列）', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'set_range', sheetId: SHEET_ID, start: 'G5', values: [['=1+1'], ['=2*2']] },
      ],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    const summary = [
      changeLine(bundle.initial, 'G5', { value: null, formula: '=1+1' }),
      changeLine(bundle.initial, 'G6', { value: null, formula: '=2*2' }),
    ].join('; ')
    expect(result.output).toBe(
      `Auto-applied 2 change(s) (undo via the side panel [Undo] button or ⌘Z): ${summary}` +
        '\nFormula results: G5 = (still computing; verify with read_cells); G6 = (still computing; verify with read_cells)',
    )
    expect(cellOf(bundle, 'G5')?.formula).toBe('=1+1')
    expect(cellOf(bundle, 'G6')?.formula).toBe('=2*2')
  })

  it('add_chart 输出含 visual id，且该 id 可直接驱动 edit_chart（R3 F12a）', async () => {
    // R2 T44/T45：add_chart 输出不含 id，模型连猜 12 个 id 触发 8 连败熔断。
    // 牙点：id 必须逐字出现在输出里，且用输出里的 id 调 edit_chart 必须成功——
    // 只断言 label 文案而不验证 id 可用，label 与真实 id 漂移时测试是绿的。
    const bundle = live()
    // fixture 的 B 列全是文本（F 列公式格在 bug-hunt 无缓存值），先播种数值——
    // add_chart 要求至少一列数值。
    const seed = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'set_cell', sheetId: SHEET_ID, address: 'B2', value: 3 },
        { op: 'set_cell', sheetId: SHEET_ID, address: 'B3', value: 5 },
      ],
      summary: 'seed',
    })
    expect(seed.isError).toBeFalsy()
    const add = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'add_chart', sheetId: SHEET_ID, chartType: 'bar', dataRange: 'A1:B3', title: 'T' },
      ],
      summary: 't',
    })
    expect(add.isError).toBeFalsy()
    const visuals =
      bundle.adapter.getSnapshot().sheets.find((s) => s.id === SHEET_ID)?.visuals ?? []
    expect(visuals).toHaveLength(1)
    const id = visuals[0]!.id // length 1 asserted above
    expect(id).toMatch(/^demo-chart-/)
    expect(add.output).toBe(
      `Auto-applied 1 change(s) (undo via the side panel [Undo] button or ⌘Z): ` +
        `Insert bar chart from A1:B3 "T" (visual id: ${id} — use this id with edit_chart / delete_visual)`,
    )
    const edit = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'edit_chart', chartPath: id, chartType: 'line' }],
      summary: 't',
    })
    expect(edit.isError).toBeFalsy()
    expect(edit.output).toContain('Auto-applied 1 change(s)')
  })

  it('getActiveSheetInfo 列出 live 新增图表（R3 F12b：sim 对齐真实 app 的 charts 派生）', async () => {
    const bundle = live()
    expect(bundle.deps.getActiveSheetInfo().charts ?? []).toHaveLength(0)
    const seed = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'set_cell', sheetId: SHEET_ID, address: 'B2', value: 7 }],
      summary: 'seed',
    })
    expect(seed.isError).toBeFalsy()
    const add = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'add_chart', sheetId: SHEET_ID, chartType: 'pie', dataRange: 'A1:B3', title: 'PieT' },
      ],
      summary: 't',
    })
    expect(add.isError).toBeFalsy()
    // 真实 app 的 get_workbook_context 会把它列进 "Charts in the workbook"；
    // harness 曾漏掉这段派生，sim 里模型永远看不到刚建的图（R2 T44 四次
    // get_workbook_context 全落空）。
    expect(bundle.deps.getActiveSheetInfo().charts).toEqual([
      // types 是存储态 chartTypes（CHART_EDIT_TYPES: pie → 'pieChart'），不是编辑 token 'pie'——
      // 真实 app 与 harness 都原样展示存储态，断言跟存储态走。
      { path: expect.stringMatching(/^demo-chart-/), title: 'PieT', types: 'pieChart', sheetId: SHEET_ID },
    ])
  })

  it('copy_range 把 "Sheet!range" 塞进 source 被拒且文案指向 sourceSheetId（R3 F14）', async () => {
    const bundle = live({ extraSheets: [{ id: 'sheet-2', name: '统计' }] })
    // R2 T24 的真实输入形状：sheet 名前缀混在 source 里——整批被拒，
    // 但错误文案必须告诉模型 sheet 该放哪个参数。
    const bad = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'copy_range', sheetId: 'sheet-2', source: `${SHEET_NAME}!A1:F5`, target: 'A1' },
      ],
      summary: 't',
    })
    expect(bad.isError).toBeTruthy()
    expect(bad.output).toContain('sheet-name prefix')
    expect(bad.output).toContain('sourceSheetId')
    // 修正形态（sourceSheetId + 纯 range）通过并真的写进目标 sheet。
    const good = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'copy_range', sheetId: 'sheet-2', sourceSheetId: SHEET_ID, source: 'A1:F5', target: 'A1' },
      ],
      summary: 't',
    })
    expect(good.isError).toBeFalsy()
    expect(cellOf(bundle, 'A1', 'sheet-2')?.value).toBe(bundle.initial['A1']?.value)
  })
})

// ---------------------------------------------------------------------------
// 9. 写闭环：bulk op（fill/copy/move + 回读角格）
// ---------------------------------------------------------------------------

describe('propose_operations 写闭环（bulk op）', () => {
  it('fill_range 数值源平铺 5 格：快照 5 格 + 回读首尾角', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'fill_range', sheetId: SHEET_ID, source: 'A1', target: 'A100:A104' },
      ],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    const src = bundle.initial['A1']
    for (const row of [100, 101, 102, 103, 104]) {
      expect(cellOf(bundle, `A${row}`)).toEqual({ value: src?.value })
    }
    // fill 目标角格回读（cellChanges 无公式 → 只读 first/last）
    expect(result.output).toContain(
      `\nFormula results: A100 = ${src?.value}; A104 = ${src?.value}`,
    )
    expect(result.output).not.toContain('still computing')
  })

  it('fill_range 公式源：相对引用偏移落盘 + 回读行重复（cellChanges 与角格各一次）', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'fill_range', sheetId: SHEET_ID, source: 'F3', target: 'F100:F101' },
      ],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expect(cellOf(bundle, 'F100')?.formula).toBe('=EVALUATE(E100)')
    expect(cellOf(bundle, 'F101')?.formula).toBe('=EVALUATE(E101)')
    // demo 引擎不计算 → 4 行 still computing；F100/F101 各出现两次
    // （plan.cellChanges 与 fill 目标角格各 push 一次——观感重复，见 BUG 清单观察项）
    const still = '(still computing; verify with read_cells)'
    expect(result.output).toContain(
      `\nFormula results: F100 = ${still}; F101 = ${still}; F100 = ${still}; F101 = ${still}`,
    )
  })

  it('copy_range 块复制（1 行 × 2 列）：A200/B200 落盘 + 回读首尾角', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'copy_range', sheetId: SHEET_ID, source: 'A1:B1', target: 'A200' },
      ],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expect(cellOf(bundle, 'A200')?.value).toBe(bundle.initial['A1']?.value)
    expect(cellOf(bundle, 'B200')?.value).toBe(bundle.initial['B1']?.value)
    // 源只有 1 行 → 201 行不被触碰（保留原值）
    expect(cellOf(bundle, 'A201')?.value).toBe(bundle.initial['A201']?.value)
    expect(cellOf(bundle, 'B201')?.value).toBe(bundle.initial['B201']?.value)
    // 回读角：first=A200，last=B200
    expect(result.output).toContain(
      `\nFormula results: A200 = ${bundle.initial['A1']?.value}; B200 = ${bundle.initial['B1']?.value}`,
    )
  })

  it('copy_range filterColumn：匹配行紧凑落静态值（无公式）+ 只回读锚点', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        {
          op: 'copy_range',
          sheetId: SHEET_ID,
          source: 'B2:B5',
          target: 'G200',
          filterColumn: 'B',
          filterValues: ['m²'],
        },
      ],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    // B2=项 B3=m³ B4=m² B5=m³ → 只有 B4 匹配 → 1 行落到 G200
    expect(cellOf(bundle, 'G200')).toEqual({ value: 'm²' })
    expect(cellOf(bundle, 'G201')).toBeUndefined()
    expect(result.output).toContain('\nFormula results: G200 = m²')
    // 静态值写入不产生 formula 回读以外的行
    expect(result.output).not.toContain('G200 = (still')
  })

  it('move_range 脱糖（copy+clear）：目标写入 + 源清空 + 回读角格', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'move_range', sheetId: SHEET_ID, source: 'A30:B30', target: 'A40' },
      ],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expect(cellOf(bundle, 'A40')?.value).toBe(bundle.initial['A30']?.value)
    expect(cellOf(bundle, 'B40')?.value).toBe(bundle.initial['B30']?.value)
    expect(cellOf(bundle, 'A30')?.value ?? null).toBeNull()
    expect(cellOf(bundle, 'B30')?.value ?? null).toBeNull()
    expect(result.output).toContain(
      `\nFormula results: A40 = ${bundle.initial['A30']?.value}; B40 = ${bundle.initial['B30']?.value}`,
    )
  })
})

// ---------------------------------------------------------------------------
// 10. 写闭环：structural op
// ---------------------------------------------------------------------------

describe('propose_operations 写闭环（structural）', () => {
  it('insert_rows：行下移 + 公式相对引用同步偏移 + 标签', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'insert_rows', sheetId: SHEET_ID, row: 100, count: 1 }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expectAutoApplied(result.output, 1, 'Insert 1 row before row 100')
    // 新第 100 行为空；原 A100 值下移到 A101
    expect(cellOf(bundle, 'A100')?.value ?? null).toBeNull()
    expect(cellOf(bundle, 'A101')?.value).toBe(bundle.initial['A100']?.value)
    // F100（原 F99 公式）/ F101（原 F100 =EVALUATE(E100) → 偏移为 =EVALUATE(E101)）
    expect(cellOf(bundle, 'F101')?.formula).toBe('=EVALUATE(E101)')
  })

  it('delete_rows 删除被引用行 → #REF! 警告 + 公式文本 #REF!', async () => {
    const bundle = live()
    // 批 1：造一个引用 E3 的公式（structural 与 cell 分批复提）
    const first = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'set_formula', sheetId: SHEET_ID, address: 'G5', formula: '=E3' }],
      summary: 'ref E3',
    })
    expect(first.isError).toBeFalsy()
    // 批 2：删第 3 行 → G5→G4 公式变 #REF!
    const second = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'delete_rows', sheetId: SHEET_ID, row: 3, count: 1 }],
      summary: 'del row 3',
    })
    expect(second.isError).toBeFalsy()
    expect(second.output).toBe(
      `Auto-applied 1 change(s) (undo via the side panel [Undo] button or ⌘Z): Delete row 3` +
        `\nNote: Formulas reference deleted targets and will break (#REF!): ${SHEET_NAME}!G4`,
    )
    expect(cellOf(bundle, 'G4')?.formula).toBe('=#REF!')
    // F 列同列引用行随行移位（F100 原 =EVALUATE(E100) → 行 99 的 =EVALUATE(E99)）
    expect(cellOf(bundle, 'F99')?.formula).toBe('=EVALUATE(E99)')
  })

  it('delete_rows 多行标签 Delete rows a–b', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'delete_rows', sheetId: SHEET_ID, row: 10, count: 2 }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expectAutoApplied(result.output, 1, 'Delete rows 10–11')
  })

  it('insert_cols 标签 + 数据不动（G 为末列）', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'insert_cols', sheetId: SHEET_ID, column: 'H', count: 1 }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expectAutoApplied(result.output, 1, 'Insert 1 column before column H')
    expect(cellOf(bundle, 'G100')?.value ?? null).toBe(bundle.initial['G100']?.value ?? null)
  })

  it('delete_cols C:D：E/F/G 左移 + F 列公式列引用同步', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'delete_cols', sheetId: SHEET_ID, column: 'C', count: 2 }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expectAutoApplied(result.output, 1, 'Delete columns C–D')
    // E(5)→C(3)、F(6)→D(4)：原 F100 =EVALUATE(E100) 现居 D100，引用 E→C
    expect(cellOf(bundle, 'C100')?.value).toBe(bundle.initial['E100']?.value)
    expect(cellOf(bundle, 'D100')?.formula).toBe('=EVALUATE(C100)')
  })

  it('add_sheet 标签 + 快照多一 sheet', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'add_sheet', name: '汇总' }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expectAutoApplied(result.output, 1, 'Add sheet "汇总"')
    const sheets = bundle.adapter.getSnapshot().sheets
    expect(sheets.length).toBe(2)
    expect(sheets[1]?.name).toBe('汇总')
  })

  it('rename_sheet 与 set_cell 混批：opCount 双算 + 快照改名', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 7 },
        { op: 'rename_sheet', sheetId: SHEET_ID, name: '工程量' },
      ],
      summary: 'mix',
    })
    expect(result.isError).toBeFalsy()
    expectAutoApplied(
      result.output,
      2,
      `${changeLine(bundle.initial, 'G5', { value: 7 })} | sheet ${SHEET_NAME} → 工程量`,
    )
    expect(bundle.adapter.getSnapshot().sheets[0]?.name).toBe('工程量')
  })

  it('delete_sheet 跨 sheet 公式引用 → danglers 警告（带引号 sheet 名）', async () => {
    const bundle = live({
      extraSheets: [
        { id: 'sheet-2', name: 'Sheet2', cells: { A1: { formula: `='${SHEET_NAME}'!E3` } } },
      ],
    })
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'delete_sheet', sheetId: 'sheet-2' }],
      summary: 't',
      confirmDestructive: true, // F11：delete_sheet 无条件破坏性，确认后走 danglers 校验
    })
    // 删 sheet-2（引用方）而非被引用方 → 无 danglers
    expect(result.isError).toBeFalsy()
    expect(result.output).not.toContain('#REF!')
    const again = live({
      extraSheets: [
        { id: 'sheet-2', name: 'Sheet2', cells: { A1: { formula: `='${SHEET_NAME}'!E3` } } },
      ],
    })
    const result2 = await runTool(again.deps, 'propose_operations', {
      operations: [{ op: 'delete_sheet', sheetId: SHEET_ID }],
      summary: 't',
      confirmDestructive: true,
    })
    expect(result2.isError).toBeFalsy()
    expect(result2.output).toContain(
      `\nNote: Formulas reference deleted targets and will break (#REF!): Sheet2!A1`,
    )
  })
})

// ---------------------------------------------------------------------------
// 11. 回读细节（上限 / 跨 sheet 前缀 / notices）
// ---------------------------------------------------------------------------

describe('propose_operations 回读细节', () => {
  it('回读 10 格上限：12 个公式 → 10 行 + …2 more', async () => {
    const bundle = live()
    const operations = Array.from({ length: 12 }, (_, i) => ({
      op: 'set_formula' as const,
      sheetId: SHEET_ID,
      address: `A${100 + i}`,
      formula: '=1+1',
    }))
    const result = await runTool(bundle.deps, 'propose_operations', { operations, summary: 't' })
    expect(result.isError).toBeFalsy()
    const lines = result.output.split('\nFormula results: ')[1]
    expect(lines).toContain('A109 = (still computing; verify with read_cells)')
    expect(lines).not.toContain('A110 =')
    expect(lines).toContain('; …2 more formula cells')
  })

  it('cellChanges 超 20 个：summary …N more cells（36 格动态逐字）', async () => {
    const bundle = live()
    const values = Array.from({ length: 6 }, () => [1, 2, 3, 4, 5, 6])
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'set_range', sheetId: SHEET_ID, start: 'A100', values }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    const cellPart: string[] = []
    for (let r = 0; r < 6; r += 1) {
      for (let c = 0; c < 6; c += 1) {
        const address = ['A', 'B', 'C', 'D', 'E', 'F'][c]! + String(100 + r)
        cellPart.push(changeLine(bundle.initial, address, { value: values[r]![c]! }))
      }
    }
    const shown = cellPart.slice(0, 20).join('; ') + '; …16 more cells'
    expectAutoApplied(result.output, 36, shown)
  })

  it('跨 sheet 回读加 sheet 名前缀', async () => {
    const bundle = live({
      extraSheets: [{ id: 'sheet-2', name: 'Sheet2' }],
    })
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'set_formula', sheetId: SHEET_ID, address: 'G5', formula: '=1' },
        { op: 'set_formula', sheetId: 'sheet-2', address: 'A1', formula: '=2' },
      ],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    const still = '(still computing; verify with read_cells)'
    expect(result.output).toContain(
      `\nFormula results: ${SHEET_NAME}!G5 = ${still}; Sheet2!A1 = ${still}`,
    )
  })

  it('apply notices 进 Note 行', async () => {
    const { deps } = live({ failApply: () => ({ ok: true, notices: ['frozen to values'] }) })
    const result = await runTool(deps, 'propose_operations', {
      operations: [{ op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 1 }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('\nNote: frozen to values')
  })

  it('plan warnings 与 apply notices 合并（; 连接）', async () => {
    // failApply 只在 structural 批生效：第一批（建引用）必须真实落盘，
    // 否则第二批 plan 时工作簿里没有公式，#REF! warning 不会产生
    const bundle = live({
      failApply: (plan) =>
        plan.structuralChanges.length > 0 ? { ok: true, notices: ['n1'] } : null,
    })
    const first = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'set_formula', sheetId: SHEET_ID, address: 'G5', formula: '=E3' }],
      summary: 'ref',
    })
    expect(first.isError).toBeFalsy()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'delete_rows', sheetId: SHEET_ID, row: 3, count: 1 }],
      summary: 'del',
    })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain(
      `\nNote: Formulas reference deleted targets and will break (#REF!): ${SHEET_NAME}!G4; n1`,
    )
  })
})

// ---------------------------------------------------------------------------
// 12. lazy 写闭环（stub apply：校验链 + 回读渲染）
// ---------------------------------------------------------------------------

describe('propose_operations lazy 写闭环（stub apply）', () => {
  it('lazy set_formula：回读显示文件缓存值（stub 边界：readCells 不随写入变化）', async () => {
    const { deps } = lazyDeps()
    const result = await runTool(deps, 'propose_operations', {
      operations: [{ op: 'set_formula', sheetId: SHEET_ID, address: 'F100', formula: '=1/0' }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain(UNDO_HINT)
    // F100 文件缓存值（oracle：fixture.cached['F100']）
    expect(result.output).toContain('Formula results: F100 = 0.4054368')
  })

  it('lazy set_cell：回读无公式格 → 无 Formula results 段', async () => {
    const { deps } = lazyDeps()
    const result = await runTool(deps, 'propose_operations', {
      operations: [{ op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 9 }],
      summary: 't',
    })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain(UNDO_HINT)
    expect(result.output).not.toContain('Formula results')
  })
})

// ---------------------------------------------------------------------------
// 13. 破坏性守卫（R3 F11：超规模破坏须先确认，代码层强制而非提示词）
// ---------------------------------------------------------------------------

describe('propose_operations 破坏性守卫（R3 F11）', () => {
  it('clear_range A1:F2000（R1 T46 形状，12,000 格）无确认 → 被拒，文案含格数 + confirmDestructive 引导，工作簿未动', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'clear_range', sheetId: SHEET_ID, range: 'A1:F2000' }],
      summary: '把所有数据都删掉，我要重新开始',
    })
    expect(result.isError).toBe(true)
    expect(result.output).toContain('destructive at scale')
    expect(result.output).toContain('12000 cells')
    expect(result.output).toContain('confirmDestructive: true')
    expect(result.output).toContain('Nothing was changed')
    // 守卫在 proposeOperations 之前拦截 → 一格未清
    expect(cellOf(bundle, 'A1')?.value ?? null).toBe(bundle.initial['A1']?.value ?? null)
  })

  it('delete_rows count=31（R2 T28 形状，低于阈值）无确认放行——有界编辑不是擦除', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'delete_rows', sheetId: SHEET_ID, row: 50, count: 31 }],
      summary: '删 31 行',
    })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('Auto-applied')
  })

  it('delete_rows count=100（恰好阈值）无确认被拒、confirmDestructive:true 后真删', async () => {
    const bundle = live()
    const refused = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'delete_rows', sheetId: SHEET_ID, row: 50, count: 100 }],
      summary: '删 100 行',
    })
    expect(refused.isError).toBe(true)
    expect(refused.output).toContain('100 rows')
    expect(refused.output).toContain('confirmDestructive: true')
    const confirmed = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'delete_rows', sheetId: SHEET_ID, row: 50, count: 100 }],
      summary: '删 100 行',
      confirmDestructive: true,
    })
    expect(confirmed.isError).toBeFalsy()
    expect(confirmed.output).toContain('Auto-applied')
  })

  it('clear_range 60 格（低于阈值）无确认放行', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'clear_range', sheetId: SHEET_ID, range: 'A1:F10' }],
      summary: '清一小块',
    })
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('Auto-applied')
  })

  it('delete_sheet 无条件须确认（sparse 视图判断不了表是否为空）；确认后 sheet 真消失', async () => {
    const bundle = live({ extraSheets: [{ id: 'sheet-2', name: 'Temp' }] })
    const refused = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'delete_sheet', sheetId: 'sheet-2' }],
      summary: '删临时表',
    })
    expect(refused.isError).toBe(true)
    expect(refused.output).toContain('delete_sheet "sheet-2"')
    expect(refused.output).toContain('confirmDestructive: true')
    const confirmed = await runTool(bundle.deps, 'propose_operations', {
      operations: [{ op: 'delete_sheet', sheetId: 'sheet-2' }],
      summary: '删临时表',
      confirmDestructive: true,
    })
    expect(confirmed.isError).toBeFalsy()
    expect(bundle.adapter.getSnapshot().sheets.map((s) => s.id)).not.toContain('sheet-2')
  })

  it('混合批：普通 set_cell + 超阈值 clear_range → 整批被拒（all-or-nothing，不部分执行）', async () => {
    const bundle = live()
    const result = await runTool(bundle.deps, 'propose_operations', {
      operations: [
        { op: 'set_cell', sheetId: SHEET_ID, address: 'G5', value: 99 },
        { op: 'clear_range', sheetId: SHEET_ID, range: 'A1:F2000' },
      ],
      summary: '改一格并清空大表',
    })
    expect(result.isError).toBe(true)
    expect(result.output).toContain('destructive at scale')
    // 守卫在批级校验前拦截 → 连 set_cell 那格也没写
    expect(cellOf(bundle, 'G5')?.value ?? null).toBe(bundle.initial['G5']?.value ?? null)
  })
})
