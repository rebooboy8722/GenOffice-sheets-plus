import { expect, test } from '@playwright/test'
import { copyFile, mkdir, mkdtemp, readdir, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { ARTIFACTS_DIR, buildGridWorkbook, launchShell, waitForPageWithUrl } from './helpers'

/**
 * Measures what the prompt changes (base.md workflow rules, the
 * `get_workbook_context` / `load_guide` tool descriptions) actually cost in
 * model round-trips, against the real local endpoint instead of a stub.
 *
 * The claim being tested is that the overview the app already appends to every
 * user message makes a `get_workbook_context` turn pure waste, and that a
 * near-unconditional "load the guide first" rule burns a whole extra turn plus
 * 3.7–9 KB of permanent history. Both are claims about **model behaviour**, so
 * no unit test can settle them — the only evidence is turn counts from a real
 * run, which is what the AI trace (Phase 0.3) was wired up to provide.
 *
 * One line per provider turn lands in GENOFFICE_AI_TRACE_DIR, so the run is
 * measured from the trace rather than from the UI: turn count, the tools each
 * turn called, per-turn wall time and token counts.
 *
 * Deliberately a measurement, not an assertion about speed: it writes a JSON
 * summary to `e2e/artifacts/turn-count-<label>.json`. Compare that file between
 * builds (patched prompts vs `git checkout v0.12.29 -- <prompt files>`), with
 * repeats — a single stochastic model run proves nothing. The assertions that
 * are here guard the *setup* (see buildTaskWorkbook).
 *
 * Needs the local model endpoint; skips when it is unreachable so the suite
 * stays runnable offline.
 */

const BASE_URL = process.env.GENOFFICE_E2E_AI_URL ?? 'http://127.0.0.1:8080/v1'
const MODEL = process.env.GENOFFICE_E2E_AI_MODEL ?? 'Qwen3.6-27B-UD-Q5_K_XL.gguf'
const INSTRUCTION = 'Format column A as a percentage (values like 0.15 should show as 15%).'
/** Labels the saved summary so repeated runs and the two arms of the comparison
 * (patched prompts vs `git checkout v0.12.29 -- <prompt files>`) stay tellable
 * apart: `GENOFFICE_E2E_TURN_LABEL=patched-1 npx playwright test …`. */
const LABEL = process.env.GENOFFICE_E2E_TURN_LABEL ?? new Date().toISOString().replace(/[:.]/g, '')
/**
 * Phase 0.2: the same task with the model's thinking turned off at the endpoint.
 * `bodyExtras` rides into the request body (custom-profile.ts → registry.ts →
 * openai-compatible.ts) and llama.cpp honours `chat_template_kwargs`.
 *
 * Worth measuring rather than eyeballing: whether thinking *feels* worth its
 * cost is the user's call, but what it costs is not — thinking tokens are
 * decoded before any tool call appears, so they are pure added latency in front
 * of the user's first sign of progress.
 */
const THINK = process.env.GENOFFICE_E2E_THINK === 'off' ? 'off' : 'on'

/** A turn's trace line, as written by packages/ai-provider/src/trace.ts. */
interface TraceLine {
  durationMs: number
  /** The provider-neutral transcript trace.ts records: a user turn carries
   * `text`, a tool turn carries `results`. */
  messages: Array<{
    role: string
    text?: string
    results?: Array<{ toolName?: string; output?: unknown; isError?: boolean }>
  }>
  response: {
    text?: string
    reasoning?: string
    toolCalls: Array<{ name: string }>
    stopReason?: string
    error?: string
    usage?: { inputTokens?: number; outputTokens?: number }
  }
}

async function endpointReachable(): Promise<boolean> {
  try {
    const response = await fetch(`${BASE_URL.replace(/\/v1$/, '')}/v1/models`, {
      signal: AbortSignal.timeout(5_000),
    })
    return response.ok
  } catch {
    return false
  }
}

async function readTrace(dir: string): Promise<TraceLine[]> {
  const files = (await readdir(dir).catch(() => [] as string[])).filter((f) => f.endsWith('.jsonl'))
  const lines: TraceLine[] = []
  for (const file of files) {
    const raw = await readFile(join(dir, file), 'utf8').catch(() => '')
    for (const line of raw.split('\n')) {
      if (!line.trim()) continue
      try {
        lines.push(JSON.parse(line) as TraceLine)
      } catch {
        /* a partially written line; the poll retries */
      }
    }
  }
  return lines
}

/**
 * The run is over when the last turn asked for no tools (the loop's terminal
 * turn) and the trace has been quiet — model-agnostic, unlike waiting on
 * specific prose.
 */
async function waitForRunEnd(dir: string, timeoutMs = 300_000): Promise<TraceLine[]> {
  const deadline = Date.now() + timeoutMs
  let lastCount = 0
  let lastChange = Date.now()
  for (;;) {
    const lines = await readTrace(dir)
    if (lines.length !== lastCount) {
      lastCount = lines.length
      lastChange = Date.now()
    }
    const last = lines.at(-1)
    if (last && last.response.toolCalls.length === 0 && Date.now() - lastChange > 3_000) return lines
    if (Date.now() > deadline) return lines
    await new Promise((r) => setTimeout(r, 1_000))
  }
}

test('measure model round-trips for one formatting task', async () => {
  test.skip(!(await endpointReachable()), `model endpoint unreachable: ${BASE_URL}`)

  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-turns-'))
  const workbook = join(scratch, 'book.xlsx')
  // A header label above the numbers: "format column A" is ambiguous about the
  // header row, and a model that has to guess spends a turn asking. The values
  // go into the file, not the grid — see buildGridWorkbook.
  await writeFile(
    workbook,
    await buildGridWorkbook([['share'], [0.15], [0.4], [0.07], [0.62], [0.28]]),
  )
  const traceDir = join(scratch, 'trace')
  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-e2e-'))

  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = {
    apiKey: 'local',
    model: MODEL,
    baseUrl: BASE_URL,
    ...(THINK === 'off'
      ? { profile: { bodyExtras: { chat_template_kwargs: { enable_thinking: false } } } }
      : {}),
  }
  await writeFile(join(userDataDir, 'ai-settings.json'), JSON.stringify(resolveCustomEndpoints(settings)))

  const launched = await launchShell({
    userDataDir,
    onboardingSeen: true,
    openFile: workbook,
    videoDir: 'turn-count',
    env: { GENOFFICE_AI_TRACE_DIR: traceDir },
  })
  try {
    const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
    await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    await sheets.waitForTimeout(2_000)

    const composer = sheets.getByPlaceholder(/Ask AI to analyze/)
    await composer.click()
    await composer.fill(INSTRUCTION)
    await sheets.getByRole('button', { name: 'Send', exact: true }).click()

    const lines = await waitForRunEnd(traceDir)
    // The last request carries the whole conversation, so it is where the tool
    // outputs and the overview the model was actually shown can be read back.
    const history = lines.at(-1)?.messages ?? []
    const toolOutputs = history.flatMap((m) => (m.results ?? []).map((r) => String(r.output ?? '')))
    const overviewShown = history.find((m) => m.role === 'user')?.text ?? ''
    const summary = {
      baseUrl: BASE_URL,
      model: MODEL,
      instruction: INSTRUCTION,
      think: THINK,
      turns: lines.length,
      totalDurationMs: lines.reduce((sum, l) => sum + l.durationMs, 0),
      // Field names are the provider-neutral ones from AiTokenUsage, as written
      // by trace.ts — not the OpenAI wire names.
      inputTokens: lines.map((l) => l.response.usage?.inputTokens ?? 0),
      outputTokens: lines.map((l) => l.response.usage?.outputTokens ?? 0),
      reasoningChars: lines.map((l) => l.response.reasoning?.length ?? 0),
      toolsPerTurn: lines.map((l) => l.response.toolCalls.map((c) => c.name)),
      errors: lines.map((l) => l.response.error).filter(Boolean),
      /** Turns lost to a batch the schema refused. A rejection whose message
       * could have explained itself costs a whole round trip, which is what
       * this counts (see packages/xlsx-gateway/tests/range-error-wording.test.ts). */
      rejectedBatches: toolOutputs.filter((o) => o.includes('Rejected — none of')).length,
      overviewShown,
    }
    console.log('TURN_SUMMARY>>>', JSON.stringify(summary, null, 1), '<<<END')
    await mkdir(ARTIFACTS_DIR, { recursive: true })
    await writeFile(
      join(ARTIFACTS_DIR, `turn-count-${LABEL}.json`),
      JSON.stringify({ label: LABEL, ...summary }, null, 2),
    )
    await sheets.screenshot({ path: `/tmp/turn-count-${LABEL}.png` })
    // Asserted last, so a failing run still prints and saves what the model was
    // shown. Guards the measurement itself: if the file's six rows did not
    // reach the overview, the run measured a different task than the one asked.
    expect(overviewShown).toContain('A1:A6')
  } finally {
    await launched.app.close().catch(() => {})
    // Keep the raw trace beside the summary: the per-turn messages are the
    // evidence behind the count (what the model was shown, what it asked for).
    for (const file of (await readdir(traceDir).catch(() => [] as string[])).filter((f) =>
      f.endsWith('.jsonl'),
    )) {
      const destDir = join(ARTIFACTS_DIR, 'trace')
      await mkdir(destDir, { recursive: true })
      await copyFile(join(traceDir, file), join(destDir, file))
    }
  }
})
