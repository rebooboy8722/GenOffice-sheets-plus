/**
 * F19（R5 triage）：harness 替身的 evaluateFormula(formula, sheetId) 接收了
 * sheetId 但从未使用——裸引用恒按硬编码 SHEET_ID（主表）解析。R5 T10 模型对
 * sheet-2 跑 =COUNTA(A1:F2300) 拿到 13,072（主表值），同范围 aggregate_range
 * 测得 4,675，模型烧 215s / 40K rchars 调和矛盾才在 T11 答对。生产代码
 * （tools.ts evaluate_formula → parseReadSheetId）正确按*请求的*表解析裸引用；
 * 本测试把替身钉回同一语义：带 sheetId 数第二张表，不带保持主表默认。
 * 牙试验：主表 fixture A1:A10 恰 10 个非空格，第二张表写 3 格——修复前此处
 * 返回 10（主表值）必红，修复后返回 3。
 */
import { describe, expect, it } from 'vitest'

import { demoDeps, SHEET_ID } from './harness'

describe('demoDeps: F19 evaluateFormula honors the requested sheet', () => {
  it('bare refs resolve to the requested sheet, not the hardcoded main sheet', async () => {
    const bundle = demoDeps({ live: true })
    const { deps } = bundle
    const adapter = bundle.adapter
    expect(adapter).toBeDefined()
    const a = adapter!

    // 建第二张表，写 3 个已知非空格
    a.apply(
      a.plan({
        dslVersion: 1,
        transactionId: 'f19-add',
        baseRevision: a.getSnapshot().revision,
        summary: 'add second sheet',
        operations: [{ op: 'add_sheet', name: '第二表', rows: 10, columns: 5 }],
      }),
    )
    const second = a.getSnapshot().sheets.find((s) => s.name === '第二表')
    expect(second).toBeDefined()
    const secondId = second!.id
    expect(secondId).not.toBe(SHEET_ID)

    a.apply(
      a.plan({
        dslVersion: 1,
        transactionId: 'f19-write',
        baseRevision: a.getSnapshot().revision,
        summary: 'write 3 cells on second sheet',
        operations: [
          { op: 'set_cell', sheetId: secondId, address: 'A1', value: 'x' },
          { op: 'set_cell', sheetId: secondId, address: 'A2', value: 'y' },
          { op: 'set_cell', sheetId: secondId, address: 'A4', value: 'z' },
        ],
      }),
    )

    // F19 核心：同范围带 sheetId=第二表，必须数第二张表（3）而非主表（10）
    const result = await deps.evaluateFormula?.('=COUNTA(A1:A10)', secondId)
    expect(result).toBe(3)

    // 不带 sheetId 保持原有默认（裸引用按主表解析）
    const fallback = await deps.evaluateFormula?.('=COUNTA(A1:A10)')
    expect(fallback).toBe(10)
  }, 30000)
})
