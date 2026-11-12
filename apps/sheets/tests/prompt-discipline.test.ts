/**
 * 提示词收敛纪律回归（思考预算耗尽三层方案的预防层）：base.md 的 General discipline
 * 末尾有两条铁律级 bullet（用户 2026-11-10 改写，fc5da179；018e18d3 梳理后"铁律级
 * 只留末两条"）。本测试钉住**存在与位置**，防未来提示词重构时误删/误移。
 *
 * 注意：文件内容断言删句必红（平凡为真），模型行为的测量在 convo-sim A/B
 * （reasoning 字符中位数 / 模型调用数 / 正确率）。
 */
import { describe, expect, it } from 'vitest'

import basePrompt from '../src/renderer/ai/prompts/base.md?raw'

describe('base.md 收敛纪律 bullet', () => {
  const discipline = basePrompt.split('# General discipline')[1] ?? ''

  it('大任务分轮小步：大文件处理必须分步处理', () => {
    expect(discipline).toContain('大文件处理必须分步处理')
  })

  it('思考保持紧凑：思考只工作方法和步骤', () => {
    expect(discipline).toContain('不要在思考里处理数据')
  })

  it('公式被拒不换参重试（R6 T19 基线 10 次无效 MATCH 调用，R7 B 段对照）', () => {
    expect(discipline).toContain("Don't reshuffle a rejected formula")
  })

  it('公式形态被拒不换形态再发整批（R16 43/20 + R17 41/13 两轮 40+/20 档判定非波动，R18 A 段对照）', () => {
    expect(discipline).toContain("Don't re-batch a rejected formula")
  })

  it('破坏性操作要确认必须同回合提交登记（F23：R8 T26–T27 确认 UX 死锁，R9 A 段对照）', () => {
    expect(discipline).toContain('so the safety gate registers the refusal')
  })

  it('scratch sheet 删不掉要报告残留（F25：R9 T12 Scratch 死公式残留未报告，L52 措辞与门控冲突）', () => {
    expect(discipline).toContain('report the leftover scratch sheet')
  })

  it('结构性删除前后必须重读 extent（F26：R9 T31 继承历史 extent 删错范围 → 8 格残留 + 跨回合错误前提继承）', () => {
    expect(discipline).toContain('never reuse an extent read earlier')
  })

  it('残留 scratch sheet 不许承诺"稍后自动清理"（F28：R10 T12 删除被门控拦后"稍后会自动清理"虚假承诺）', () => {
    expect(discipline).toContain('do not say it will be cleaned up later')
  })

  it('两条都在 # General discipline 标记之后、且分步在前思考在后', () => {
    const marker = basePrompt.indexOf('# General discipline')
    const step = basePrompt.indexOf('大文件处理必须分步处理')
    const think = basePrompt.indexOf('不要在思考里处理数据')
    expect(marker).toBeGreaterThanOrEqual(0)
    expect(step).toBeGreaterThan(marker)
    expect(think).toBeGreaterThan(marker)
    expect(think).toBeGreaterThan(step)
  })
})
