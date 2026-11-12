/**
 * F29（R13 triage）：harness 替身的 getActiveSheetInfo 曾只回 sheets:[{id,name}]，
 * 丢了 rows/columns——buildWorkbookContext 的 dims() 拿到 undefined 直接吞掉
 * "data extent about N rows" 与 "Active sheet data area: A1:G…" 两行，模型无权威
 * 行 extent 只能猜（R13 幻觉主表 2900 行，实 2313：T14 计数对但 extent 错 /
 * T25 删 2801-2900 越界 no-op / T47 报 2800 行）。真 app demo 分支
 * （workbook-readers.ts L399-412）从 cells 算 extent 报 rows/columns，替身必须
 * 跟上。本测试钉住：初始 extent = fixture（2313×7）、buildWorkbookContext 报
 * "data extent" 行、删行后 extent 收缩（live，非冻结）。
 */
import { describe, expect, it } from 'vitest'

import { buildWorkbookContext } from '../../src/renderer/ai/tools'
import { COLS, demoDeps, ROWS, SHEET_ID } from './harness'

describe('demoDeps: F29 live sheet extent (rows/columns) reported', () => {
  it('初始 extent = fixture 尺寸，buildWorkbookContext 报 data extent 行', () => {
    const { deps } = demoDeps({ live: true })
    const info = deps.getActiveSheetInfo()
    const main = info.sheets.find((s) => s.id === SHEET_ID)
    expect(main?.rows).toBe(ROWS) // 2313
    expect(main?.columns).toBe(COLS) // 7
    const ctx = buildWorkbookContext(deps)
    expect(ctx).toContain(`data extent about ${ROWS} rows`)
    expect(ctx).toContain(`Active sheet data area: A1:G${ROWS}`)
  })

  it('delete_rows 后 extent 收缩（live 重算，非冻结初始值）', () => {
    const bundle = demoDeps({ live: true })
    const { deps, adapter } = bundle
    const a = adapter!
    const before = deps.getActiveSheetInfo().sheets.find((s) => s.id === SHEET_ID)!.rows!
    // 删主表最后 100 行（2313 → 2213）；delete_rows 字段是 row（非 startRow）
    a.apply(
      a.plan({
        dslVersion: 1,
        transactionId: 't1',
        baseRevision: a.getSnapshot().revision,
        summary: 'delete last 100 rows',
        operations: [{ op: 'delete_rows', sheetId: SHEET_ID, row: ROWS - 99, count: 100 }],
      }),
    )
    const after = deps.getActiveSheetInfo().sheets.find((s) => s.id === SHEET_ID)!.rows!
    expect(after).toBe(before - 100)
  }, 30000)
})
