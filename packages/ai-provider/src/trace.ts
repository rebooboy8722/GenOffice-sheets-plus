/**
 * Opt-in model-turn trace for building fine-tuning corpora and for field
 * diagnosis. The dir comes from AiSettings.logDir (user-visible setting,
 * wins) or the GENOFFICE_AI_TRACE_DIR env var; with either set, every ai:stream
 * request appends one JSON line (full request as sent to the provider, full
 * response as streamed back, plus per-turn diagnostics) to
 * <dir>/ai-trace-<yyyymmdd>-<pid>.jsonl. Off by default; never written from
 * the renderer; the provider config (API keys) is not part of the record.
 */
import { appendFileSync, mkdirSync } from 'node:fs'
import { join } from 'node:path'
import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import type { AiTokenUsage } from './types'

export interface AiTraceRequest {
  requestId: string
  sessionId?: string | undefined
  /** which window/app issued it (e.g. the sender's URL) */
  source?: string | undefined
  provider: string
  model: string
  maxTokens: number
  system: string
  messages: AgentMessage[]
  tools: AgentToolDef[]
  /** the no-think switch was on for this request (AiSettings.nothink) */
  nothink?: boolean | undefined
}

export interface AiTraceRecorder {
  onDelta(text: string): void
  onReasoningDelta(text: string): void
  onToolCall(call: AgentToolCall): void
  onStopReason(reason: string): void
  onUsage(usage: AiTokenUsage): void
  /** writes the line; `error` for a failed turn, 'aborted' when the user stopped it */
  finish(error?: string): void
}

export function aiTraceDir(env: NodeJS.ProcessEnv = process.env): string | null {
  const dir = env.GENOFFICE_AI_TRACE_DIR?.trim()
  return dir ? dir : null
}

/** settings.logDir (the user-visible setting) wins over the env var. */
function resolveTraceDir(
  opts: { dir?: string | undefined },
  env: NodeJS.ProcessEnv,
): string | null {
  const fromSettings = opts.dir?.trim()
  return fromSettings || aiTraceDir(env)
}

/**
 * Prefix of the loop's truncation-continuation notes (loop.ts
 * TRUNCATION_CONTINUATION_NOTE / _AFTER_TEXT). The loop injects one as a
 * user message after a thinking-exhausted turn, so when it is the LAST user
 * message this request itself is an auto-continuation — the key fact for
 * diagnosing why the continuation did or did not fire on a real machine.
 */
const CONTINUATION_NOTE_PREFIX =
  '[System] Your previous reply was cut off by the output length limit'

function traceFile(dir: string, now: Date): string {
  const d = now.toISOString().slice(0, 10).replace(/-/g, '')
  return join(dir, `ai-trace-${d}-${process.pid}.jsonl`)
}

export function startAiTrace(
  request: AiTraceRequest,
  env: NodeJS.ProcessEnv = process.env,
  opts: { dir?: string | undefined } = {},
): AiTraceRecorder | null {
  const dir = resolveTraceDir(opts, env)
  if (!dir) return null
  const startedAt = new Date()
  let text = ''
  let reasoning = ''
  const toolCalls: AgentToolCall[] = []
  let stopReason: string | undefined
  let usage: AiTokenUsage | undefined
  let done = false
  return {
    onDelta: (t) => {
      text += t
    },
    onReasoningDelta: (t) => {
      reasoning += t
    },
    onToolCall: (call) => {
      toolCalls.push(call)
    },
    onStopReason: (reason) => {
      stopReason = reason
    },
    onUsage: (u) => {
      usage = u
    },
    finish: (error) => {
      if (done) return
      done = true
      // Diagnostic rollup: the fields that decide whether the loop's
      // thinking-exhaustion auto-continuation could fire (stopReason,
      // reasoningChars, toolCallCount) plus whether THIS request already IS
      // such a continuation. The blind-spot signature of a no-op r8fix is
      // stopReason=max_tokens + reasoningChars=0 with the next request
      // carrying no continuation note.
      const lastUser = [...request.messages].reverse().find((m) => m.role === 'user')
      const line = JSON.stringify({
        ts: startedAt.toISOString(),
        durationMs: Date.now() - startedAt.getTime(),
        ...request,
        response: {
          text,
          reasoning,
          toolCalls,
          stopReason,
          usage,
          error,
          diagnostics: {
            reasoningChars: reasoning.length,
            textChars: text.length,
            toolCallCount: toolCalls.length,
            nothink: request.nothink === true,
            isContinuation: lastUser?.text.startsWith(CONTINUATION_NOTE_PREFIX) ?? false,
          },
        },
      })
      try {
        mkdirSync(dir, { recursive: true })
        appendFileSync(traceFile(dir, startedAt), line + '\n')
      } catch (err) {
        console.warn('[ai-trace] write failed:', err instanceof Error ? err.message : err)
      }
    },
  }
}
