/**
 * F18（R4 triage）：harness 替身的 getActiveSheetInfo 曾把 knownAddresses
 * 固化为构造时的初始 fixture 键表——live 分支只更新 revision/sheets/charts，
 * 键表恒为初始值。R4 全程（T1 rev=0 起）上下文恒显 "first 200 of 13266"，
 * T45 模型诚实照抄 stale 数字（终态真实非空 12,840，且主表从未被清）。
 * 本测试钉住"键表必须跟随 live 快照"：clear 后计数下降、新写格后上升、
 * 被清格不残留（与 F17 读端过滤语义一致）。
 */
import { describe, expect, it } from 'vitest'

import { demoDeps, SHEET_ID } from './harness'

describe('demoDeps: F18 live knownAddresses follows the live snapshot', () => {
  it('clear shrinks the count, new writes grow it (no frozen fixture key list)', () => {
    const bundle = demoDeps({ live: true })
    const { deps } = bundle
    const adapter = bundle.adapter
    expect(adapter).toBeDefined()
    const a = adapter!

    const initial = deps.getActiveSheetInfo()
    const initialCount = initial.knownAddresses.length
    expect(initialCount).toBeGreaterThan(0)

    // 清一个区域（980 格 < F16 门控阈值 5000，直放；走 adapter.plan/apply
    // 数据层路径，该路径的空 entry 行为由 xlsx-gateway 的
    // in-memory-workbook-clear.test.ts 锚定）
    a.apply(
      a.plan({
        dslVersion: 1,
        transactionId: 't1',
        baseRevision: a.getSnapshot().revision,
        summary: 'clear A3:F100',
        operations: [{ op: 'clear_range', sheetId: SHEET_ID, range: 'A3:F100' }],
      }),
    )
    const afterClear = deps.getActiveSheetInfo()
    expect(afterClear.revision ?? 0).toBeGreaterThan(initial.revision ?? 0)
    expect(afterClear.knownAddresses.length).toBeLessThan(initialCount)
    // 被清格不得残留（"non-empty" 语义：adapter 留的 {value:null} 键被过滤）
    expect(afterClear.knownAddresses).not.toContain('A3')

    // 新写 2 格 → 计数恰增 2（证明是 live 重算，不是另一个静态表）
    a.apply(
      a.plan({
        dslVersion: 1,
        transactionId: 't2',
        baseRevision: afterClear.revision,
        summary: 'write H1:H2',
        operations: [
          { op: 'set_cell', sheetId: SHEET_ID, address: 'H1', value: 1 },
          { op: 'set_cell', sheetId: SHEET_ID, address: 'H2', value: 2 },
        ],
      }),
    )
    const afterWrite = deps.getActiveSheetInfo()
    expect(afterWrite.knownAddresses.length).toBe(afterClear.knownAddresses.length + 2)
    expect(afterWrite.knownAddresses).toContain('H1')
    expect(afterWrite.knownAddresses).toContain('H2')
  }, 30000)
})
