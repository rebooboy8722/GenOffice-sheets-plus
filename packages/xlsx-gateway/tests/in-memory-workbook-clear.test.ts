/**
 * F17（R4 triage）：demo 内存 adapter 的 clear_range 把被清的非空格**原地改写**
 * 为 `{ value: null }` entry 并保留 key（replaceCell 纯赋值、从不删 key；小范围
 * 还会被展开成逐格 set_cell(null)，连原本无 key 的空格都新建空 entry），
 * 只有 delete_rows/delete_cols 经 shiftAddressedRecord 真正丢弃 entry。
 * 读端（sheets 的 getActiveSheetInfo.knownAddresses）依赖这个事实做空 entry
 * 过滤——本测试锚定该契约：将来若 adapter 改为删除空 key，这里会红，
 * 提醒读端过滤仍要保留（行为无害但需知情）。
 */
import { describe, expect, it } from 'vitest'

import { InMemoryWorkbookAdapter } from '../src/domain/in-memory-workbook'

function seedAdapter(): InMemoryWorkbookAdapter {
  return new InMemoryWorkbookAdapter({
    revision: 0,
    sheets: [
      {
        id: 's1',
        name: 'Data',
        cells: { A1: { value: 'h' }, A2: { value: 1 }, B2: { value: 2 }, B3: { value: 3 } },
      },
    ],
  })
}

describe('in-memory workbook: destructive op cell-map maintenance (F17)', () => {
  it('clear_range keeps cleared cells in the map as {value:null} — key count does not shrink', () => {
    const adapter = seedAdapter()
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 't1',
      baseRevision: 0,
      summary: 'clear',
      operations: [{ op: 'clear_range', sheetId: 's1', range: 'A1:B2' }],
    })
    adapter.apply(plan)
    const sheet = adapter.getSnapshot().sheets[0]!
    // 小范围 clear 被展开成逐格 set_cell(null)：连原本无 key 的 B1 也被创建为
    // 空 entry——map 键数清表后不降反可增（读端过滤因此必须按内容判空）
    expect(Object.keys(sheet.cells).sort()).toEqual(['A1', 'A2', 'B1', 'B2', 'B3'])
    for (const address of ['A1', 'A2', 'B1', 'B2']) {
      expect(sheet.cells[address]!.value).toBeNull()
    }
    // 未清区域不动
    expect(sheet.cells.B3!.value).toBe(3)
  })

  it('delete_rows discards the deleted rows’ entries (the asymmetry clear_range lacks)', () => {
    const adapter = seedAdapter()
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 't1',
      baseRevision: 0,
      summary: 'delete row 2',
      operations: [{ op: 'delete_rows', sheetId: 's1', row: 2, count: 1 }],
    })
    adapter.apply(plan)
    const sheet = adapter.getSnapshot().sheets[0]!
    // 第 2 行（A2/B2）的 entry 被丢弃，第 3 行（B3=3）上移为第 2 行
    expect(Object.keys(sheet.cells).sort()).toEqual(['A1', 'B2'])
    expect(sheet.cells.B2!.value).toBe(3)
  })
})
