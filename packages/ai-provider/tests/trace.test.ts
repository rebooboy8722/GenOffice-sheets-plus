import { mkdtempSync, readdirSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { startAiTrace } from '../src/trace'

const dirs: string[] = []
afterEach(() => {
  for (const d of dirs) rmSync(d, { recursive: true, force: true })
  dirs.length = 0
})

const request = {
  requestId: 'r1',
  provider: 'genspark',
  model: 'claude-opus-5-5',
  maxTokens: 32768,
  system: 'sys',
  messages: [{ role: 'user' as const, text: 'hi' }],
  tools: [],
}

describe('startAiTrace', () => {
  it('is off without GENOFFICE_AI_TRACE_DIR', () => {
    expect(startAiTrace(request, {})).toBeNull()
  })

  it('appends one JSON line per turn with the streamed response', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ai-trace-'))
    dirs.push(dir)
    const env = { GENOFFICE_AI_TRACE_DIR: dir }
    const a = startAiTrace(request, env)!
    a.onReasoningDelta('think ')
    a.onDelta('Hel')
    a.onDelta('lo')
    a.onToolCall({ id: 't1', name: 'read_blocks', input: { from: 0 } })
    a.onStopReason('tool_use')
    a.onUsage({ inputTokens: 10, outputTokens: 5 })
    a.finish()
    a.finish()
    const b = startAiTrace({ ...request, requestId: 'r2' }, env)!
    b.finish('boom')

    const files = readdirSync(dir)
    expect(files).toHaveLength(1)
    expect(files[0]).toMatch(/^ai-trace-\d{8}-\d+\.jsonl$/)
    const lines = readFileSync(join(dir, files[0]!), 'utf8')
      .trim()
      .split('\n')
      .map((l) => JSON.parse(l))
    expect(lines).toHaveLength(2)
    expect(lines[0]).toMatchObject({
      requestId: 'r1',
      model: 'claude-opus-5-5',
      system: 'sys',
      response: {
        text: 'Hello',
        reasoning: 'think ',
        toolCalls: [{ name: 'read_blocks' }],
        stopReason: 'tool_use',
        usage: { outputTokens: 5 },
      },
    })
    expect(lines[0].response.error).toBeUndefined()
    expect(lines[1]).toMatchObject({ requestId: 'r2', response: { error: 'boom' } })
  })
})

describe('startAiTrace: per-turn diagnostics (thinking-exhaustion blind spot)', () => {
  // Same prefix as loop.ts TRUNCATION_CONTINUATION_NOTE(_AFTER_TEXT)
  const NOTE_PREFIX = '[System] Your previous reply was cut off by the output length limit'

  function readLines(dir: string) {
    const file = readdirSync(dir)[0]
    if (!file) return []
    return readFileSync(join(dir, file), 'utf8').trim().split('\n').map((l) => JSON.parse(l))
  }

  it('rolls up the fields that decide whether auto-continuation can fire', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ai-trace-'))
    dirs.push(dir)
    const a = startAiTrace(request, { GENOFFICE_AI_TRACE_DIR: dir })!
    a.onReasoningDelta('thinking a lot ')
    a.onDelta('partial')
    a.onToolCall({ id: 't1', name: 'read_range', input: { range: 'A1:B2' } })
    a.onStopReason('max_tokens')
    a.finish()
    const [line] = readLines(dir)
    expect(line.response.diagnostics).toEqual({
      reasoningChars: 15,
      textChars: 7,
      toolCallCount: 1,
      nothink: false,
      isContinuation: false,
    })
    expect(line.nothink).toBeUndefined()
  })

  it('flags a request whose last user message is a truncation-continuation note', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ai-trace-'))
    dirs.push(dir)
    const a = startAiTrace(
      {
        ...request,
        messages: [
          { role: 'user' as const, text: 'big task' },
          { role: 'assistant' as const, text: '' },
          { role: 'user' as const, text: `${NOTE_PREFIX} after thinking, before any text…` },
        ],
      },
      { GENOFFICE_AI_TRACE_DIR: dir },
    )!
    a.finish()
    expect(readLines(dir)[0].response.diagnostics.isContinuation).toBe(true)
  })

  it('records the nothink switch on the line', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ai-trace-'))
    dirs.push(dir)
    const a = startAiTrace({ ...request, nothink: true }, { GENOFFICE_AI_TRACE_DIR: dir })!
    a.finish()
    const [line] = readLines(dir)
    expect(line.nothink).toBe(true)
    expect(line.response.diagnostics.nothink).toBe(true)
  })

  it('writes to the settings logDir even when the env var points elsewhere', () => {
    const fromEnv = mkdtempSync(join(tmpdir(), 'ai-trace-env-'))
    const fromSettings = mkdtempSync(join(tmpdir(), 'ai-trace-settings-'))
    dirs.push(fromEnv, fromSettings)
    const a = startAiTrace(request, { GENOFFICE_AI_TRACE_DIR: fromEnv }, { dir: fromSettings })!
    a.finish()
    expect(readdirSync(fromSettings)).toHaveLength(1)
    expect(readdirSync(fromEnv)).toEqual([])
  })

  it('a blank settings logDir falls back to the env var', () => {
    const fromEnv = mkdtempSync(join(tmpdir(), 'ai-trace-env-'))
    dirs.push(fromEnv)
    const a = startAiTrace(request, { GENOFFICE_AI_TRACE_DIR: fromEnv }, { dir: '   ' })!
    a.finish()
    expect(readdirSync(fromEnv)).toHaveLength(1)
  })
})
