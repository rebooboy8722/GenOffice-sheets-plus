/**
 * F16 破坏性确认门控（R3 triage：F11 的 confirmDestructive 标志被模型抢跑绕过——
 * T35 首次提交即带标志直通 delete_sheet；T8 被拒后同回合立即带标志重试直通
 * delete_rows）。门控规则（机械、无语义判断）：破坏性批必须先被拒一次，
 * 且放行只在"更早的用户回合"里带标志的再提交才生效（onRunStart 每用户回合
 * tick 一次 runSeq）；同回合带标志重试与首次带标志提交一律再拒。
 *
 * 牙试验对应关系：
 * - "同回合带标志重试仍被拒" 在 F11 旧代码下会放行（绿→红的反证）
 * - "上一回合被拒 + 本回合带标志 → 放行" 挡住过度收紧（永远拒绝的回归）
 * - "首次带标志被拒" 对应 T35；"不同批次独立" 挡住 A 批确认顺带放行 B 批
 * 全黑盒走 createWorkbookSkill（含 onRunStart + gateWorkbookWrite 完整链路）。
 */
import { afterEach, describe, expect, it } from 'vitest'
import type { AgentToolCall, ToolExecution } from '@genoffice/agent-core'

import { createWorkbookSkill } from '../../src/renderer/ai/workbook-skill'
import { stableBatchKey } from '../../src/renderer/ai/destructive-guard'
import { demoDeps, restoreGlobals, SHEET_ID } from './harness'

afterEach(() => {
  restoreGlobals()
})

function setup() {
  const bundle = demoDeps({ live: true })
  if (!bundle.adapter) throw new Error('live bundle incomplete')
  const skill = createWorkbookSkill(bundle.deps)
  const snapshot = () => JSON.stringify(bundle.adapter!.getSnapshot())
  /** 一次用户回合（loop.run 的边界，真 app / sim 里由 loop 触发） */
  const run = () => {
    skill.onRunStart?.()
  }
  const call = (input: Record<string, unknown>): Promise<ToolExecution> => {
    const c: AgentToolCall = { id: 't', name: 'propose_operations', input }
    return Promise.resolve(skill.executeTool(c))
  }
  /** R3 T8 形状：删 100 行（恰好阈值） */
  const del100 = (extra: Record<string, unknown> = {}) =>
    call({
      operations: [{ op: 'delete_rows', sheetId: SHEET_ID, row: 50, count: 100 }],
      summary: '删 100 行',
      ...extra,
    })
  /** 大范围清表（12,000 格） */
  const clearBig = (extra: Record<string, unknown> = {}) =>
    call({
      operations: [{ op: 'clear_range', sheetId: SHEET_ID, range: 'A1:F2000' }],
      summary: '清空数据区',
      ...extra,
    })
  return { run, call, del100, clearBig, snapshot }
}

describe('F16 destructive confirmation gate (session-scoped)', () => {
  it('T8 形状：无标志被拒 + 同回合带标志重试 → 仍被拒（旧 F11 代码此处放行）', async () => {
    const s = setup()
    s.run()
    const refused = await s.del100()
    expect(refused.isError).toBe(true)
    expect(refused.output).toContain('destructive at scale')
    expect(refused.output).toContain('Nothing was changed')

    const retried = await s.del100({ confirmDestructive: true })
    expect(retried.isError).toBe(true)
    expect(retried.output).toContain('same turn')
    expect(retried.output).toContain('Nothing was changed')
  })

  it('T36 形状：上一用户回合被拒 + 本回合带标志（用户已确认）→ 放行且真删', async () => {
    const s = setup()
    const before = s.snapshot()
    s.run()
    expect((await s.del100()).isError).toBe(true)
    s.run() // 用户："确认，删"
    const applied = await s.del100({ confirmDestructive: true })
    expect(applied.isError).toBeFalsy()
    expect(applied.output).toContain('Auto-applied')
    expect(s.snapshot()).not.toBe(before)
  })

  it('T35 形状：首次提交即带标志（抢跑）→ 被拒并记录；用户确认回合再带标志 → 放行', async () => {
    const s = setup()
    const before = s.snapshot()
    s.run()
    const preemptive = await s.del100({ confirmDestructive: true })
    expect(preemptive.isError).toBe(true)
    expect(preemptive.output).toContain('before it was ever refused')
    expect(s.snapshot()).toBe(before)

    s.run() // 用户看到模型转述的确认请求后确认
    const applied = await s.del100({ confirmDestructive: true })
    expect(applied.isError).toBeFalsy()
    expect(applied.output).toContain('Auto-applied')
  })

  it('不同破坏性批独立记账：A 批确认放行不替 B 批免确认', async () => {
    const s = setup()
    s.run()
    expect((await s.del100()).isError).toBe(true)
    s.run()
    // A 批（上一回被拒）带标志放行
    expect((await s.del100({ confirmDestructive: true })).isError).toBeFalsy()
    // B 批从未被拒：本回合带标志也放行不了
    const b = await s.clearBig({ confirmDestructive: true })
    expect(b.isError).toBe(true)
    expect(b.output).toContain('before it was ever refused')
  })

  it('低于阈值的编辑不进门控（带标志直通，且状态不被污染）', async () => {
    const s = setup()
    s.run()
    const small = await s.call({
      operations: [{ op: 'clear_range', sheetId: SHEET_ID, range: 'A1:F10' }],
      summary: '清一小块',
      confirmDestructive: true,
    })
    expect(small.isError).toBeFalsy()
    // 大破坏性批的确认流程不受影响
    const big = await s.del100({ confirmDestructive: true })
    expect(big.isError).toBe(true)
  })

  it('stableBatchKey：同内容不同属性顺序同 key；内容不同必不同 key', () => {
    expect(stableBatchKey([{ op: 'delete_rows', sheetId: 's', row: 5, count: 100 }])).toBe(
      stableBatchKey([{ count: 100, row: 5, sheetId: 's', op: 'delete_rows' }]),
    )
    expect(stableBatchKey([{ a: 1, b: [2, 3] }])).toBe(stableBatchKey([{ b: [2, 3], a: 1 }]))
    expect(stableBatchKey([{ a: 1 }])).not.toBe(stableBatchKey([{ a: 2 }]))
  })
})
