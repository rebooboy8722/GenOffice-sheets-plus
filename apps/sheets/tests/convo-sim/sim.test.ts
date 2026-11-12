/**
 * 对话模拟器驱动（主任务：模拟真实用户逐轮对话改造真实大表，监控全流程。
 * 跑完一轮→triage→修复→下一轮验证上轮修复+新方向，共 20 轮）。
 *
 * 保真度：与 app 同一栈——AgentLoop（守卫/压缩/截断继续全活）+ createWorkbookSkill
 * （真实 base.md 提示词 + WORKBOOK_TOOLS + gateWorkbookWrite + verifySheetsResponse）
 * + streamForProvider('custom')（真实协议层，openai-compatible SSE）+ 210 vLLM 端点
 * + demoDeps({live:true})（真实 InMemoryWorkbookAdapter 写闭环，状态跨回合累积，
 * 公式格按真实 xlsx 形态种入缓存值）。
 *
 * 运行：CONVO_SIM_ROUND=1 npx vitest run tests/convo-sim/sim.test.ts
 * （不设 CONVO_SIM_ROUND 时整个套件秒级跳过，不污染常规测试。）
 *
 * 断点续跑（长任务看门配套）：每回合结束落 state.json（对话历史 + 活快照）。
 * 同轮目录已有 state.json 时自动续跑——跳过已完成回合、恢复历史与表状态；
 * 后台任务被 2 小时上限杀掉后重跑同一命令即从断点继续。
 *  - CONVO_SIM_LIMIT=N   本次启动最多执行 N 个新回合（0=全部）
 *  - CONVO_SIM_FRESH=1   清空该轮既有状态从头跑
 *
 * 产物：apps/sheets/reports/convo-sim/round-N/{events.jsonl,turns.jsonl,state.json,summary.json,config.json}
 */
import {
  appendFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterAll, beforeAll, describe, test } from 'vitest'
import {
  AgentLoop,
  type AgentLoopEvents,
  type AgentMessage,
  type AgentRunResult,
} from '@genoffice/agent-core'
import type { WorkbookSnapshot } from '@genoffice/xlsx-gateway/domain/workbook.types'

import { aiLangDirective } from '../../src/renderer/i18n/locale'
import { createWorkbookSkill } from '../../src/renderer/ai/workbook-skill'
import { demoDeps, fixture } from '../bug-hunt/harness'
import { getRoundScript } from './scripts'
import { createSimTransport } from './transport'

// ── 运行参数（env 可调）────────────────────────────────────────────
const ROUND = Number(process.env.CONVO_SIM_ROUND ?? 0)
const ENABLED = Number.isInteger(ROUND) && ROUND >= 1
const script = ENABLED ? getRoundScript(ROUND) : []
const BASE_URL = process.env.CONVO_SIM_URL ?? 'http://127.0.0.1:8080/v1'
const MODEL = process.env.CONVO_SIM_MODEL ?? 'qwen3.8-27b-quasar-v152'
/** 与应用默认一致：DEFAULT_MAX_OUTPUT_TOKENS=32768（reasoning 与正文共享该预算） */
const MAX_TOKENS = Number(process.env.CONVO_SIM_MAX_TOKENS ?? 32_768)
/** 单用户回合软上限：到点 cancel 并记 timeout，继续下一回合 */
const TURN_SOFT_TIMEOUT_MS = Number(process.env.CONVO_SIM_TURN_TIMEOUT_MS ?? 20 * 60_000)
/** 210 vLLM max_model_len */
const CONTEXT_WINDOW = 262_144
const LIMIT = Number(process.env.CONVO_SIM_LIMIT ?? 0)
const FRESH = process.env.CONVO_SIM_FRESH === '1'

const HERE = dirname(fileURLToPath(import.meta.url))
const OUT_DIR = join(HERE, '..', '..', 'reports', 'convo-sim', `round-${ROUND}`)
const EVENTS_PATH = join(OUT_DIR, 'events.jsonl')
const TURNS_PATH = join(OUT_DIR, 'turns.jsonl')
const STATE_PATH = join(OUT_DIR, 'state.json')

/** 断点：已完成回合数 + 对话历史 + 活快照 */
interface SimState {
  done: number
  at: string
  history: AgentMessage[]
  snapshot: WorkbookSnapshot
}

let resumedDone = 0
if (ENABLED) {
  if (FRESH) {
    for (const f of [STATE_PATH, TURNS_PATH, EVENTS_PATH, join(OUT_DIR, 'summary.json')]) {
      if (existsSync(f)) unlinkSync(f)
    }
  }
  if (existsSync(STATE_PATH)) {
    try {
      resumedDone = Math.max(0, Math.floor((JSON.parse(readFileSync(STATE_PATH, 'utf8')) as SimState).done))
    } catch {
      resumedDone = 0
    }
  }
}
const stopAt = LIMIT > 0 ? Math.min(script.length, resumedDone + LIMIT) : script.length

function logEvent(entry: Record<string, unknown>): void {
  appendFileSync(EVENTS_PATH, JSON.stringify({ ts: new Date().toISOString(), ...entry }) + '\n')
}

// ── 回合记录 ─────────────────────────────────────────────────────
interface ToolCallLog {
  name: string
  inputPreview: string
  ms: number
  isError: boolean
  mutated: boolean
  outputLen: number
  outputPreview: string
}
interface TurnRecord {
  t: number
  instruction: string
  status: 'ok' | 'error' | 'timeout'
  ms: number
  modelTurns: number
  /** 本回合全部模型调用的 reasoning 字符总和（对照实验指标 A） */
  reasoningChars: number
  toolCalls: ToolCallLog[]
  truncated: boolean
  turnLimit: boolean
  usage: { input: number; output: number }
  textPreview: string
  /** F8: 完整终局回答；F20: 无条件落盘（原 >2000 字符阈值在 R5 T32 失灵——
   *  10 行 top-10 表格 ≈1.8K 字符没过阈值，事后无法复核余 9 行；50 回合 ×
   *  数 KB 相对 state.json 的 MB 级可忽略，可复核性优先于省空间） */
  textFull?: string
  error?: string
}

type Outcome = { kind: 'done'; result: AgentRunResult } | { kind: 'error'; error: string }
interface ActiveTurn {
  record: TurnRecord
  startedAt: number
  settle: (out: Outcome) => void
  settled: boolean
  softFired: boolean
}

function readAllTurns(): TurnRecord[] {
  try {
    return readFileSync(TURNS_PATH, 'utf8')
      .split('\n')
      .filter(Boolean)
      .map((l) => JSON.parse(l) as TurnRecord)
  } catch {
    return []
  }
}

function summarize(records: TurnRecord[]) {
  return {
    round: ROUND,
    done: records.length,
    stopAt,
    ok: records.filter((r) => r.status === 'ok').length,
    err: records.filter((r) => r.status === 'error').length,
    timeout: records.filter((r) => r.status === 'timeout').length,
    totalMs: records.reduce((s, r) => s + r.ms, 0),
    toolCalls: records.reduce((s, r) => s + r.toolCalls.length, 0),
    toolErrors: records.reduce((s, r) => s + r.toolCalls.filter((t) => t.isError).length, 0),
    truncatedRuns: records.filter((r) => r.truncated).length,
    turnLimitRuns: records.filter((r) => r.turnLimit).length,
    perTurn: records.map((r) => ({
      t: r.t,
      status: r.status,
      ms: Math.round(r.ms / 1000),
      turns: r.modelTurns,
      reasoning: r.reasoningChars,
      tools: r.toolCalls.length,
      toolErr: r.toolCalls.filter((t) => t.isError).length,
      trunc: r.truncated,
      limit: r.turnLimit,
    })),
  }
}

describe.skipIf(!ENABLED)(`convo-sim round ${ROUND}`, () => {
  let loop!: AgentLoop
  let adapter!: { getSnapshot(): WorkbookSnapshot }
  let active: ActiveTurn | null = null
  /** 下一回合开始时仍未落地的回合（vitest 超时强推）：迟到的 onDone/onError 路由到这里，绝不落到新回合 */
  const orphaned: ActiveTurn[] = []
  let heartbeat: ReturnType<typeof setInterval> | null = null
  let summaryLogged = false
  let aborted = false
  let badStreak = 0
  /** 当前模型调用的 reasoning 累计长度（onReasoning 每次 delta 带全文，取 length 即可） */
  let reasoningLen = 0

  /** onToolStart → onToolExecuted 的配对表（同一 user 回合内工具 id 唯一） */
  const toolFinish = new Map<
    string,
    (e: { isError?: boolean; mutated?: boolean; output: string }) => void
  >()

  const events: AgentLoopEvents<unknown> = {
    onText: (text) => {
      if (active && text) active.record.textPreview = text.slice(0, 200)
    },
    onToolStart: (call) => {
      const rec = active?.record
      if (!rec) return
      const entry: ToolCallLog = {
        name: call.name,
        inputPreview: safeJson(call.input).slice(0, 300),
        ms: 0,
        isError: false,
        mutated: false,
        outputLen: -1,
        outputPreview: '',
      }
      rec.toolCalls.push(entry)
      const startedAt = Date.now()
      toolFinish.set(call.id, (execution) => {
        entry.ms = Date.now() - startedAt
        entry.isError = !!execution.isError
        entry.mutated = !!execution.mutated
        entry.outputLen = execution.output.length
        entry.outputPreview = execution.output.slice(0, 200)
        logEvent({
          kind: 'tool',
          t: rec.t,
          name: call.name,
          isError: entry.isError,
          mutated: entry.mutated,
          ms: entry.ms,
          input: entry.inputPreview,
          outputLen: entry.outputLen,
          output: entry.outputPreview,
        })
      })
    },
    onToolExecuted: (ev) => {
      const finish = toolFinish.get(ev.call.id)
      if (finish) {
        finish(ev.execution)
        toolFinish.delete(ev.call.id)
      }
    },
    onReasoning: (reasoning) => {
      if (active) reasoningLen = reasoning.length
    },
    onTurnEnd: () => {
      if (!active) return
      active.record.modelTurns += 1
      active.record.reasoningChars += reasoningLen
      logEvent({ kind: 'model-turn', t: active.record.t, n: active.record.modelTurns, reasoning: reasoningLen })
      reasoningLen = 0
    },
    onUsage: (usage) => {
      if (!active) return
      active.record.usage.input += usage.inputTokens ?? 0
      active.record.usage.output += usage.outputTokens ?? 0
    },
    onDone: (result) => {
      // settle 目标：优先未落地的 active；否则最早的孤儿回合（防迟到 onDone 误结算新回合）
      const target = active && !active.settled ? active : (orphaned.shift() ?? null)
      if (!target) {
        logEvent({ kind: 'settle-dropped', via: 'done' })
        return
      }
      if (target !== active) logEvent({ kind: 'late-settle', t: target.record.t, via: 'done' })
      else if (orphaned.length > 0)
        logEvent({ kind: 'settle-ambiguous', t: target.record.t, orphans: orphaned.map((o) => o.record.t) })
      if (reasoningLen > 0) {
        // 最后一次模型调用（纯文本无工具）没有 onTurnEnd，在此补齐 per-call 记录
        // （F7：该次调用也是一次模型回合，modelTurns 一并 +1）
        target.record.reasoningChars += reasoningLen
        target.record.modelTurns += 1
        logEvent({ kind: 'model-turn', t: target.record.t, n: target.record.modelTurns, final: true, reasoning: reasoningLen })
        reasoningLen = 0
      }
      target.record.truncated = !!result.truncated
      target.record.turnLimit = result.turnLimit
      if (!target.record.textPreview) target.record.textPreview = (result.text ?? '').slice(0, 200)
      if (result.text) target.record.textFull = result.text
      logEvent({
        kind: 'settle',
        t: target.record.t,
        via: 'done',
        ms: Date.now() - target.startedAt,
        cancelled: result.cancelled,
      })
      target.settle({ kind: 'done', result })
    },
    onError: (error) => {
      const target = active && !active.settled ? active : (orphaned.shift() ?? null)
      if (!target) {
        logEvent({ kind: 'settle-dropped', via: 'error', error: error.slice(0, 200) })
        return
      }
      if (target !== active) logEvent({ kind: 'late-settle', t: target.record.t, via: 'error' })
      else if (orphaned.length > 0)
        logEvent({ kind: 'settle-ambiguous', t: target.record.t, orphans: orphaned.map((o) => o.record.t) })
      if (reasoningLen > 0) {
        target.record.reasoningChars += reasoningLen
        logEvent({ kind: 'model-turn', t: target.record.t, n: target.record.modelTurns + 1, final: true, reasoning: reasoningLen })
        reasoningLen = 0
      }
      target.record.error = error
      logEvent({ kind: 'settle', t: target.record.t, via: 'error', ms: Date.now() - target.startedAt, error: error.slice(0, 200) })
      target.settle({ kind: 'error', error })
    },
  }

  function safeJson(value: unknown): string {
    try {
      return JSON.stringify(value)
    } catch {
      return String(value)
    }
  }

  async function waitIdle(ms: number): Promise<void> {
    const end = Date.now() + ms
    while (loop.busy && Date.now() < end) {
      await new Promise((r) => setTimeout(r, 500))
    }
  }

  async function runTurn(t: number, instruction: string): Promise<TurnRecord> {
    const record: TurnRecord = {
      t,
      instruction,
      status: 'ok',
      ms: 0,
      modelTurns: 0,
      reasoningChars: 0,
      toolCalls: [],
      truncated: false,
      turnLimit: false,
      usage: { input: 0, output: 0 },
      textPreview: '',
    }
    const startedAt = Date.now()
    let resolveOutcome!: (out: Outcome) => void
    const outcome = new Promise<Outcome>((resolve) => {
      resolveOutcome = resolve
    })
    const turn: ActiveTurn = {
      record,
      startedAt,
      settled: false,
      softFired: false,
      settle: (out) => {
        if (turn.settled) return
        turn.settled = true
        resolveOutcome(out)
      },
    }
    if (active && !active.settled) {
      // 上一回合未落地就被 vitest 超时强推：孤儿化并记录现场
      orphaned.push(active)
      logEvent({
        kind: 'orphaned',
        t: active.record.t,
        ms: Date.now() - active.startedAt,
        busy: loop.busy,
        res: process.getActiveResourcesInfo(),
      })
    }
    active = turn
    const softTimer = setTimeout(() => {
      turn.softFired = true
      logEvent({
        kind: 'soft-timeout',
        t,
        softMs: TURN_SOFT_TIMEOUT_MS,
        busy: loop.busy,
        res: process.getActiveResourcesInfo(),
      })
      loop.cancel()
      logEvent({ kind: 'cancel-issued', t, res: process.getActiveResourcesInfo() })
    }, TURN_SOFT_TIMEOUT_MS)
    // 硬底线：cancel 后 5 分钟仍不落地就放弃等待，避免整轮挂死
    const hardTimer = setTimeout(
      () => {
        logEvent({ kind: 'hard-deadline', t, res: process.getActiveResourcesInfo() })
        turn.settle({
          kind: 'error',
          error: `hard deadline after cancel (${Math.round(TURN_SOFT_TIMEOUT_MS / 60000)}+5min)`,
        })
      },
      TURN_SOFT_TIMEOUT_MS + 5 * 60_000,
    )

    await waitIdle(60_000)
    if (loop.busy) {
      clearTimeout(softTimer)
      clearTimeout(hardTimer)
      active = null
      record.ms = Date.now() - startedAt
      record.status = 'error'
      record.error = 'loop still busy from previous turn'
      return record
    }
    loop.run(instruction)
    const out = await outcome
    clearTimeout(softTimer)
    clearTimeout(hardTimer)
    active = null
    record.ms = Date.now() - startedAt
    if (out.kind === 'error') {
      record.status = turn.softFired ? 'timeout' : 'error'
      record.error = out.error
    } else {
      record.status = turn.softFired ? 'timeout' : 'ok'
    }
    return record
  }

  beforeAll(async () => {
    mkdirSync(OUT_DIR, { recursive: true })
    // 端点探针：模型在列才开跑（离线时秒级红，不烧回合）
    const res = await fetch(`${BASE_URL}/models`, { signal: AbortSignal.timeout(20_000) })
    if (!res.ok) throw new Error(`endpoint probe HTTP ${res.status}`)
    const body = (await res.json()) as { data?: Array<{ id: string }> }
    if (!body.data?.some((m) => m.id === MODEL)) {
      throw new Error(`model ${MODEL} not listed at ${BASE_URL}`)
    }

    // 断点恢复：state.json 存在 → 恢复表状态与对话历史
    let state: SimState | null = null
    if (resumedDone > 0 && existsSync(STATE_PATH)) {
      try {
        state = JSON.parse(readFileSync(STATE_PATH, 'utf8')) as SimState
      } catch {
        state = null
      }
    }
    const bundle = demoDeps({
      live: true,
      seedFormulaValues: fixture.cached,
      ...(state ? { initialSnapshot: state.snapshot } : {}),
    })
    adapter = bundle.adapter!
    loop = new AgentLoop({
      transport: createSimTransport({
        baseUrl: BASE_URL,
        model: MODEL,
        maxTokens: MAX_TOKENS,
        contextWindow: CONTEXT_WINDOW,
      }),
      skill: createWorkbookSkill(bundle.deps),
      events,
      // 与应用一致：语言指令（默认 zh）
      systemSuffix: () => aiLangDirective(),
    })
    if (state && state.history.length > 0) loop.restore(state.history)
    writeFileSync(
      join(OUT_DIR, 'config.json'),
      JSON.stringify(
        {
          round: ROUND,
          baseUrl: BASE_URL,
          model: MODEL,
          maxTokens: MAX_TOKENS,
          contextWindow: CONTEXT_WINDOW,
          resumedFrom: resumedDone,
          stopAt,
          softMs: TURN_SOFT_TIMEOUT_MS,
          launchedAt: new Date().toISOString(),
        },
        null,
        2,
      ),
    )
    console.log(
      `[sim R${ROUND}] ${resumedDone > 0 ? `resume after turn ${resumedDone}` : 'fresh'} stopAt=${stopAt} url=${BASE_URL} model=${MODEL} softMs=${Math.round(TURN_SOFT_TIMEOUT_MS / 60000)}m`,
    )
    // 挂死取证：60s 心跳——进程活性 + 活动异步资源（区分 hung-socket / lost-promise / 事件循环阻塞 / 进程被挂起）
    heartbeat = setInterval(() => {
      logEvent({
        kind: 'heartbeat',
        t: active?.record.t ?? null,
        settled: active?.settled ?? null,
        softFired: active?.softFired ?? null,
        busy: loop.busy,
        elapsedMs: active ? Date.now() - active.startedAt : null,
        orphans: orphaned.map((o) => o.record.t),
        res: process.getActiveResourcesInfo(),
      })
    }, 60_000)
    // 进程退出兜底（exit/SIGINT/SIGTERM 可捕获；SIGKILL 捕获不到——最后一条心跳即死亡上界）
    const deathDump = (kind: string) => {
      try {
        logEvent({
          kind,
          t: active?.record.t ?? null,
          settled: active?.settled ?? null,
          busy: loop.busy,
          res: process.getActiveResourcesInfo(),
        })
      } catch {
        /* 退出中的尽力而为 */
      }
    }
    process.once('exit', () => deathDump('process-exit'))
    process.once('SIGINT', () => {
      deathDump('sigint')
      process.exit(130)
    })
    process.once('SIGTERM', () => {
      deathDump('sigterm')
      process.exit(143)
    })
  }, 120_000)

  for (let i = 1; i <= script.length; i++) {
    const msg = script[i - 1] ?? ''
    test(`T${i} ${msg.slice(0, 24)}`, {
      timeout: TURN_SOFT_TIMEOUT_MS + 10 * 60_000,
      skip: i <= resumedDone || i > stopAt,
    }, async () => {
      if (aborted) return
      const rec = await runTurn(i, msg)
      appendFileSync(TURNS_PATH, JSON.stringify(rec) + '\n')
      // 断点落盘（历史 + 活快照）：被杀后重跑从此处继续
      writeFileSync(
        STATE_PATH,
        JSON.stringify({
          done: i,
          at: new Date().toISOString(),
          history: [...loop.messages],
          snapshot: adapter.getSnapshot(),
        } satisfies SimState),
      )
      const toolErr = rec.toolCalls.filter((x) => x.isError).length
      const mutated = rec.toolCalls.some((x) => x.mutated)
      console.log(
        `[sim R${ROUND} T${String(i).padStart(2, '0')}] ${rec.status} ${Math.round(rec.ms / 1000)}s turns=${rec.modelTurns} reasoning=${rec.reasoningChars} tools=${rec.toolCalls.length} toolErr=${toolErr} mut=${mutated} trunc=${rec.truncated} limit=${rec.turnLimit} tok=${rec.usage.input}/${rec.usage.output} | ${rec.textPreview.replace(/\n/g, ' ').slice(0, 60)}${rec.error ? ` | ERR: ${rec.error.slice(0, 150)}` : ''}`,
      )
      // 连续失败熔断：基础设施挂了不再烧后面的回合
      badStreak = rec.status === 'ok' ? 0 : badStreak + 1
      if (badStreak >= 4 || (resumedDone === 0 && i === 3 && readAllTurns().every((r) => r.status !== 'ok'))) {
        aborted = true
        console.log(`[sim R${ROUND}] ABORT: 连续 ${badStreak} 回合非 ok，停止后续回合（写 summary 后结束）`)
      }
    })
  }

  afterAll(async () => {
    if (heartbeat) clearInterval(heartbeat)
    if (summaryLogged) return
    const records = readAllTurns()
    writeFileSync(join(OUT_DIR, 'summary.json'), JSON.stringify(summarize(records), null, 2))
    const s = summarize(records)
    console.log(
      `[sim R${ROUND}] done ok=${s.ok} err=${s.err} timeout=${s.timeout} tools=${s.toolCalls} toolErr=${s.toolErrors} trunc=${s.truncatedRuns} totalMs=${s.totalMs} aborted=${aborted}`,
    )
  }, 60_000)
})
