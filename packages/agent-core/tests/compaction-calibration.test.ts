import { describe, expect, it, vi } from 'vitest'
import {
  AgentLoop,
  type AgentSkill,
  type AgentStreamCallbacks,
  type AgentStreamRequest,
  type AgentTransport,
} from '../src'

const flush = () => new Promise<void>((resolve) => setTimeout(resolve, 0))

const skill: AgentSkill = {
  id: 'test',
  systemPrompt: 'Test system prompt',
  tools: [{ name: 'read', description: 'Read', inputSchema: { type: 'object' } }],
  executeTool: () => ({ output: 'ok', summary: 'Read completed' }),
}

/** 'Old answer ' × 1819 = 20009 bytes: above the 8192-window calibrated budget (18432), below the 256KB shipped default */
const BIG_ANSWER = 'Old answer '.repeat(1819)

type TransportCall = { request: AgentStreamRequest; callbacks: AgentStreamCallbacks }

function makeTransport(window?: number) {
  const calls: TransportCall[] = []
  const transport: AgentTransport = {
    stream(request, callbacks) {
      calls.push({ request, callbacks })
      return { cancel: () => undefined }
    },
    ...(window !== undefined ? { contextWindow: () => window } : {}),
  }
  return { transport, calls }
}

/**
 * Build the foldable shape compaction needs: an old big exchange plus a recent
 * one. The third run's beginRun is where the budget check fires.
 */
async function buildHistory(transport: AgentTransport, calls: TransportCall[]) {
  const loop = new AgentLoop({
    transport,
    skill,
    compaction: { disableLlmSummary: true },
  })
  loop.run('Old conversation instruction')
  await flush()
  calls[0]!.callbacks.onDelta(BIG_ANSWER)
  calls[0]!.callbacks.onDone()
  loop.run('Recent question')
  await flush()
  calls[1]!.callbacks.onDelta('Recent answer')
  calls[1]!.callbacks.onDone()
  return loop
}

describe('compaction budget calibration from the transport window', () => {
  it('a small published window compacts where the shipped defaults would not', async () => {
    // 8192 tokens → maxBytes floor(8192 × 3 × 0.75) = 18432; the 20009-byte
    // old answer pushes the history over it, but stays far under the 256KB default
    for (const window of [8192, undefined]) {
      const { transport, calls } = makeTransport(window)
      const loop = await buildHistory(transport, calls)
      loop.run('Continue the conversation')
      await flush()
      if (window === 8192) {
        const first = calls[2]!.request.messages[0]!
        expect(first).toEqual({
          role: 'user',
          text: expect.stringContaining('[Summary of earlier conversation'),
        })
        // the mechanical digest still names the folded instruction
        expect((first as { role: 'user'; text: string }).text).toContain(
          'Old conversation instruction',
        )
      } else {
        // uncalibrated: 20066 bytes < 256KB, history stays untouched
        expect(calls[2]!.request.messages[0]).toEqual({
          role: 'user',
          text: 'Old conversation instruction',
        })
      }
    }
  })

  it('a huge published window relaxes the trigger to 75% of it (documented, not accidental)', async () => {
    // 262144 tokens (the 210 vLLM shape) → maxBytes 589824; a 300KB history is
    // over the 256KB shipped default but under the calibrated budget
    const big = 'A'.repeat(300_000)
    for (const window of [262_144, undefined]) {
      const { transport, calls } = makeTransport(window)
      const loop = new AgentLoop({
        transport,
        skill,
        compaction: { disableLlmSummary: true },
      })
      loop.run('Old conversation instruction')
      await flush()
      calls[0]!.callbacks.onDelta(big)
      calls[0]!.callbacks.onDone()
      loop.run('Recent question')
      await flush()
      calls[1]!.callbacks.onDelta('Recent answer')
      calls[1]!.callbacks.onDone()
      loop.run('Continue the conversation')
      await flush()
      expect(calls[2]!.request.messages[0]).toEqual(
        window === 262_144
          ? { role: 'user', text: 'Old conversation instruction' } // no compaction at the calibrated budget
          : { role: 'user', text: expect.stringContaining('[Summary of earlier conversation') },
      )
    }
  })

  it('an explicit compaction option outranks the window calibration', async () => {
    // 8192 window would say 18432; the explicit 300000 says otherwise
    const { transport, calls } = makeTransport(8192)
    const loop = new AgentLoop({
      transport,
      skill,
      compaction: { disableLlmSummary: true, maxBytes: 300_000 },
    })
    loop.run('Old conversation instruction')
    await flush()
    calls[0]!.callbacks.onDelta(BIG_ANSWER)
    calls[0]!.callbacks.onDone()
    loop.run('Recent question')
    await flush()
    calls[1]!.callbacks.onDelta('Recent answer')
    calls[1]!.callbacks.onDone()
    loop.run('Continue the conversation')
    await flush()
    expect(calls[2]!.request.messages[0]).toEqual({
      role: 'user',
      text: 'Old conversation instruction',
    })
  })

  it('a throwing contextWindow falls back to the shipped defaults', () => {
    // restore() compacts without a surrounding try/catch, so an uncaught throw
    // here would be a crash, not a fallback — this pins the guard down
    const { transport } = makeTransport(8192)
    const broken: AgentTransport = {
      stream: transport.stream.bind(transport),
      contextWindow: () => {
        throw new Error('broken')
      },
    }
    const loop = new AgentLoop({ transport: broken, skill })
    loop.restore([
      { role: 'user', text: 'Old conversation instruction' },
      { role: 'assistant', text: 'A'.repeat(300_000) }, // over the 256KB default
      { role: 'user', text: 'Recent question' },
      { role: 'assistant', text: 'Recent answer' },
    ])
    expect(loop.messages[0]).toEqual({
      role: 'user',
      text: expect.stringContaining('[Summary of earlier conversation'),
    })
  })
})

const OVERFLOW_400 =
  "HTTP 400: This model's maximum context length is 8192 tokens. However, your messages already contain 9500 tokens, which exceeds the model's maximum context length for this model."

function scriptedTransport(
  window: number | undefined,
  script: Array<(cb: AgentStreamCallbacks) => void>,
) {
  let i = 0
  const calls: AgentStreamRequest[] = []
  const transport: AgentTransport = {
    stream(request, callbacks) {
      calls.push(request)
      script[i++]?.(callbacks)
      return { cancel: () => undefined }
    },
    ...(window !== undefined ? { contextWindow: () => window } : {}),
  }
  return { transport, calls }
}

async function buildOverflowSetup(
  window: number | undefined,
  script: Array<(cb: AgentStreamCallbacks) => void>,
) {
  const { transport, calls } = scriptedTransport(window, [
    (cb) => {
      cb.onDelta(BIG_ANSWER)
      cb.onDone()
    },
    (cb) => {
      cb.onDelta('Recent answer')
      cb.onDone()
    },
    ...script,
  ])
  const onError = vi.fn()
  const onDone = vi.fn()
  const loop = new AgentLoop({
    transport,
    skill,
    events: { onError, onDone },
    compaction: { disableLlmSummary: true },
  })
  loop.run('Old conversation instruction')
  await flush()
  loop.run('Recent question')
  await flush()
  return { loop, calls, onError, onDone }
}

describe('400 context-overflow rescue: compact, replay once', () => {
  it('replays the turn after compaction when the provider reports a context overflow', async () => {
    const { loop, calls, onError, onDone } = await buildOverflowSetup(8192, [
      (cb) => cb.onError(OVERFLOW_400),
      (cb) => {
        cb.onDelta('Recovered')
        cb.onDone()
      },
    ])

    onError.mockClear()
    onDone.mockClear()
    loop.run('Overflowing instruction')
    await flush()
    // the replay schedules itself behind maybeCompact's promise chain
    await flush()
    await flush()

    expect(calls).toHaveLength(4)
    // the replayed request carries the compacted history, not the one that overflowed
    expect(calls[2]!.messages[0]).toEqual({
      role: 'user',
      text: expect.stringContaining('[Summary of earlier conversation'),
    })
    expect(calls[3]!.messages[0]).toEqual(calls[2]!.messages[0])
    expect(onError).not.toHaveBeenCalled()
    expect(onDone).toHaveBeenCalledWith(
      expect.objectContaining({ text: 'Recovered', cancelled: false }),
    )
    expect(loop.busy).toBe(false)
  })

  it('fails the run when a second overflow follows the replay', async () => {
    const { loop, calls, onError, onDone } = await buildOverflowSetup(8192, [
      (cb) => cb.onError(OVERFLOW_400),
      (cb) => cb.onError(OVERFLOW_400),
    ])

    onError.mockClear()
    onDone.mockClear()
    loop.run('Overflowing instruction')
    await flush()
    await flush()
    await flush()

    expect(calls).toHaveLength(4)
    expect(onError).toHaveBeenCalledTimes(1)
    expect(onError).toHaveBeenCalledWith(OVERFLOW_400)
    expect(onDone).not.toHaveBeenCalled()
    expect(loop.busy).toBe(false)
  })

  it('does not rescue a 400 that is not a context overflow', async () => {
    const { loop, calls, onError, onDone } = await buildOverflowSetup(8192, [
      (cb) => cb.onError('HTTP 400: invalid_api_key'),
    ])

    onError.mockClear()
    onDone.mockClear()
    loop.run('Another instruction')
    await flush()
    await flush()

    expect(calls).toHaveLength(3) // no replay
    expect(onError).toHaveBeenCalledTimes(1)
    expect(onError).toHaveBeenCalledWith('HTTP 400: invalid_api_key')
    expect(onDone).not.toHaveBeenCalled()
    expect(loop.busy).toBe(false)
  })

  it('does not rescue window phrasing without a 400 status marker', async () => {
    const { loop, calls, onError, onDone } = await buildOverflowSetup(8192, [
      (cb) => cb.onError('maximum context length exceeded for this model'),
    ])

    onError.mockClear()
    onDone.mockClear()
    loop.run('Another instruction')
    await flush()
    await flush()

    expect(calls).toHaveLength(3) // no replay
    expect(onError).toHaveBeenCalledTimes(1)
    expect(onError).toHaveBeenCalledWith('maximum context length exceeded for this model')
    expect(onDone).not.toHaveBeenCalled()
    expect(loop.busy).toBe(false)
  })

  it('Anthropic-style wording with a provider prefix also triggers the rescue', async () => {
    const { loop, calls, onError, onDone } = await buildOverflowSetup(8192, [
      (cb) => cb.onError('Claude HTTP 400: prompt is too long: 9500 tokens > 8192 token limit'),
      (cb) => {
        cb.onDelta('Recovered')
        cb.onDone()
      },
    ])

    onError.mockClear()
    onDone.mockClear()
    loop.run('Overflowing instruction')
    await flush()
    await flush()
    await flush()

    expect(calls).toHaveLength(4)
    expect(onDone).toHaveBeenCalledWith(
      expect.objectContaining({ text: 'Recovered' }),
    )
    expect(loop.busy).toBe(false)
  })
})

/**
 * Calibrating the budget on the tokens the provider actually billed.
 *
 * The shipped constant assumes ~3 bytes/token (prose edge: CJK). Digit-dense
 * spreadsheet content — the read_range output a sheets run lives on — measures
 * ~1.0–1.4 bytes/token against the real Qwen3.8-27B tokenizer, so a budget
 * sized on the constant holds 1.7× the real window for exactly that content:
 * the first 400 becomes a certainty and the once-per-run overflow rescue
 * cannot cover a long run's second overflow. The fix learns the ratio from
 * the last turn's billed prompt; these tests pin the calibration path and
 * its two fallbacks (no report, bogus sample).
 */

/** 450_000 bytes of digit-dense rows, shaped like read_range output */
const DENSE_TABLE = '1,234,567\t2,345.67\n'.repeat(22_500)

async function buildDenseSetup(billedRun2: number | null) {
  const { transport, calls } = scriptedTransport(262_144, [
    (cb) => {
      if (billedRun2 !== null) cb.onUsage?.({ inputTokens: 3_000 })
      cb.onDelta(DENSE_TABLE)
      cb.onDone()
    },
    (cb) => {
      if (billedRun2 !== null) cb.onUsage?.({ inputTokens: billedRun2 })
      cb.onDelta('Recent answer')
      cb.onDone()
    },
  ])
  const loop = new AgentLoop({
    transport,
    skill,
    compaction: { disableLlmSummary: true },
  })
  loop.run('Old conversation instruction')
  await flush()
  loop.run('Recent question')
  await flush()
  return { loop, calls }
}

describe('compaction budget calibrated from provider-billed tokens', () => {
  it('a dense mix the provider bills compacts where the shipped 3 B/tok constant would not', async () => {
    // run 2's history (~450KB of digits) bills 350K tokens ≈ 1.29 B/tok:
    // the shipped constant budgets 262144 × 3 × 0.75 = 589824 bytes (no
    // compaction), the learned ratio budgets ≈ 252KB → compaction on run 3.
    const { loop, calls } = await buildDenseSetup(350_000)
    loop.run('Continue the conversation')
    await flush()
    expect(calls).toHaveLength(3)
    // run 2 went out uncalibrated: run 1's history (a 28-byte instruction) is
    // under the 2KB threshold, so its 3K-token bill — mostly system prompt —
    // teaches nothing about the mix
    expect(calls[1]!.messages[0]).toEqual({
      role: 'user',
      text: 'Old conversation instruction',
    })
    // run 3 compacts on the mix the provider measured in run 2
    expect(calls[2]!.messages[0]).toEqual({
      role: 'user',
      text: expect.stringContaining('[Summary of earlier conversation'),
    })
  })

  it('keeps the shipped 3 B/tok behavior when the provider never reports usage', async () => {
    const { loop, calls } = await buildDenseSetup(null)
    loop.run('Continue the conversation')
    await flush()
    expect(calls).toHaveLength(3)
    // 450KB < 589824 (uncalibrated budget): untouched, exactly the pre-calibration behavior
    expect(calls[2]!.messages[0]).toEqual({
      role: 'user',
      text: 'Old conversation instruction',
    })
  })

  it('clamps a bogus billing sample to the safety band instead of breaking the budget', async () => {
    // 450KB of history "billed" at 10 tokens measures 45000 B/tok — outside the
    // 0.5–6 band; clamped to 6 the budget is 262144 × 6 × 0.75 = 1179648 bytes,
    // so 450KB stays under it (the band limits one bad sample; correctness then
    // rides the 400-overflow rescue)
    const { loop, calls } = await buildDenseSetup(10)
    loop.run('Continue the conversation')
    await flush()
    expect(calls).toHaveLength(3)
    expect(calls[2]!.messages[0]).toEqual({
      role: 'user',
      text: 'Old conversation instruction',
    })
  })
})
