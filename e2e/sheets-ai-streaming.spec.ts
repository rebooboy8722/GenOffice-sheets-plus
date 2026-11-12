import { expect, test } from '@playwright/test'
import { createServer, type Server } from 'node:http'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { AddressInfo } from 'node:net'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { buildGridWorkbook, launchShell, waitForPageWithUrl } from './helpers'

/**
 * Streaming behaviour in the real app: how many times an answer is committed to
 * the DOM, and whether the model's thinking shows up at all.
 *
 * Both claims are about what the renderer *does* while tokens arrive, so unit
 * tests can only cover half of each:
 *  - the coalescer (`ai/stream-coalescer.ts`) is tested against a fake scheduler;
 *    that it is wired to `onText` at all, and that frames actually run here, is
 *    this test's job.
 *  - the reasoning disclosure (`AiChatPanel`) is tested in jsdom; the chain that
 *    feeds it — main process → IPC chunk → `AgentLoopEvents.onReasoning` → panel
 *    — spans three processes and no unit test crosses it.
 *
 * The stub streams one answer as SSE the way a local server does: reasoning
 * deltas first, then text deltas. DOM mutations inside the transcript are
 * counted while it streams, with a `requestAnimationFrame` counter beside them —
 * a frame-less measurement (hidden window, stalled compositor) would otherwise
 * look like perfect throttling.
 */

/**
 * Long enough to matter. A 60-token answer is cheap to render by any method, and
 * the cost the coalescer targets grows with the answer: `onText` hands over the
 * **cumulative** text, so every commit re-parses the whole Markdown body — a
 * per-delta commit path is quadratic in length. Measuring on a short answer
 * would show no difference and prove nothing.
 */
const REASONING =
  'The user wants a total row. Reading the extent: A1:B5, so row 6 is free. ' +
  'A label there would not collide with data, and B6 can hold the sum. ' +
  'I should keep the header row untouched and write only the two new cells. '

const ANSWER = [
  '## What I changed',
  '',
  '- **A6** — the label `total`',
  '- **B6** — `=SUM(B2:B5)`, so it follows the data instead of hardcoding 10',
  '',
  '| cell | before | after |',
  '| --- | --- | --- |',
  '| A6 | (empty) | total |',
  '| B6 | (empty) | 10 |',
  '',
  'The header row and the five data rows are untouched, and the format of column B',
  'is unchanged, so the percentages still display as they did. Undo with ⌘Z.',
].join('\n')

/** Deltas the stub emits per field: enough to make a per-delta commit obvious. */
const CHUNKS = 220



function splitForStreaming(text: string, parts: number): string[] {
  const size = Math.ceil(text.length / parts)
  const out: string[] = []
  for (let i = 0; i < text.length; i += size) out.push(text.slice(i, i + size))
  return out
}

function startStreamingStub(port = 0): Promise<Server> {
  const server = createServer(async (_req, res) => {
    res.writeHead(200, {
      'content-type': 'text/event-stream',
      'cache-control': 'no-cache',
      connection: 'keep-alive',
    })
    const send = (payload: unknown) => res.write(`data: ${JSON.stringify(payload)}\n\n`)
    send({ choices: [{ index: 0, delta: { role: 'assistant', content: '' } }] })
    // Reasoning first, as a thinking model streams it: the screen sits on the
    // typing indicator for the whole of this stretch.
    // Paced at ~125 tokens/s, the local endpoint's decode speed and the case the
    // coalescer exists for: deltas arriving faster than the page paints. Unglued
    // (0) the whole answer lands in one read, React 19 batches it into a handful
    // of commits, and the test would pass whether the coalescer works or not —
    // measured: 8 mutations and 50 ms with the gap at 0, 211 and 611 ms at 8.
    const GAP = Number(process.env.GENOFFICE_E2E_STREAM_GAP_MS ?? 8)
    for (const piece of splitForStreaming(REASONING, CHUNKS)) {
      send({ choices: [{ index: 0, delta: { reasoning_content: piece } }] })
      if (GAP) await new Promise((r) => setTimeout(r, GAP))
    }
    for (const piece of splitForStreaming(ANSWER, CHUNKS)) {
      send({ choices: [{ index: 0, delta: { content: piece } }] })
      if (GAP) await new Promise((r) => setTimeout(r, GAP))
    }
    send({ choices: [{ index: 0, delta: {}, finish_reason: 'stop' }] })
    res.write('data: [DONE]\n\n')
    res.end()
  })
  return new Promise((r) => server.listen(port, '127.0.0.1', () => r(server)))
}

interface StreamProbe {
  /** DOM mutations inside the transcript, the visible cost of a commit */
  mutations: number
  /** Frames the page actually painted while streaming */
  frames: number
}

/** Cumulative main-thread time from Chromium (seconds), read across the stream. */
interface ThreadTime {
  script: number
  style: number
  layout: number
}

const TIMING_METRICS: Record<string, keyof ThreadTime> = {
  ScriptDuration: 'script',
  RecalcStyleDuration: 'style',
  LayoutDuration: 'layout',
}

async function readThreadTime(cdp: {
  send(method: 'Performance.getMetrics'): Promise<{ metrics: Array<{ name: string; value: number }> }>
}): Promise<ThreadTime> {
  const { metrics } = await cdp.send('Performance.getMetrics')
  const out: ThreadTime = { script: 0, style: 0, layout: 0 }
  for (const m of metrics) {
    const field = TIMING_METRICS[m.name]
    if (field) out[field] = m.value
  }
  return out
}

function minus(a: ThreadTime, b: ThreadTime): ThreadTime {
  return { script: a.script - b.script, style: a.style - b.style, layout: a.layout - b.layout }
}

const ms = (seconds: number): number => Math.round(seconds * 1000)

/**
 * Counts transcript DOM mutations and animation frames for the duration of the
 * run. Mutation records, not render counts: React may commit without changing
 * text, and those are exactly the commits the coalescer exists to avoid — they
 * still cost layout.
 */
async function installProbe(page: import('@playwright/test').Page): Promise<void> {
  await page.evaluate(() => {
    const host = document.querySelector('.ai-chat')
    if (!host) throw new Error('AI transcript (.ai-chat) not found')
    const probe = { mutations: 0, frames: 0 }
    ;(window as unknown as Record<string, unknown>).__streamProbe = probe
    new MutationObserver((records) => {
      probe.mutations += records.length
    }).observe(host, { childList: true, subtree: true, characterData: true })
    const tick = () => {
      probe.frames += 1
      requestAnimationFrame(tick)
    }
    requestAnimationFrame(tick)
  })
}

async function readProbe(page: import('@playwright/test').Page): Promise<StreamProbe> {
  return page.evaluate(
    () => (window as unknown as Record<string, StreamProbe>).__streamProbe,
  )
}

test('streaming: one answer costs frames, not tokens, and its thinking is shown', async () => {
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-stream-'))
  const workbook = join(scratch, 'book.xlsx')
  await writeFile(workbook, await buildGridWorkbook([['name'], ['alpha'], ['beta']]))

  const server = await startStreamingStub()
  const port = (server.address() as AddressInfo).port
  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-e2e-'))
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = {
    apiKey: 'stub',
    model: 'stub-model',
    baseUrl: `http://127.0.0.1:${port}/v1`,
  }
  await writeFile(join(userDataDir, 'ai-settings.json'), JSON.stringify(resolveCustomEndpoints(settings)))

  const launched = await launchShell({
    userDataDir,
    onboardingSeen: true,
    openFile: workbook,
    videoDir: 'streaming',
  })
  try {
    const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
    await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    await sheets.waitForTimeout(2_000)

    // Main-thread time, not just DOM writes: the cost of a commit is the render
    // work behind it (a 4400-line App tree plus a full Markdown re-parse of the
    // cumulative text), most of which changes no DOM at all. Mutation counts
    // alone understate it — measured, they separated the two arms by 1.7× where
    // scripting time separated them by far more.
    const cdp = await launched.app.context().newCDPSession(sheets)
    await cdp.send('Performance.enable')
    await installProbe(sheets)
    const before = await readThreadTime(cdp)

    const composer = sheets.getByPlaceholder(/Ask AI to analyze/)
    await composer.click()
    await composer.fill('Label the row below the data.')
    await sheets.getByRole('button', { name: 'Send', exact: true }).click()

    // A fragment from the last line: the answer is Markdown, so the whole body
    // is not one text node to match against.
    await expect(sheets.getByText('Undo with ⌘Z.', { exact: false }).first()).toBeVisible({
      timeout: 30_000,
    })
    const spent = minus(await readThreadTime(cdp), before)
    const probe = await readProbe(sheets)
    console.log(
      'STREAM_PROBE>>>',
      JSON.stringify({
        ...probe,
        deltasSent: CHUNKS * 2,
        answerChars: ANSWER.length,
        /** Commits per frame: the number the fix is supposed to keep ≤1. */
        mutationsPerFrame: probe.frames > 0 ? probe.mutations / probe.frames : null,
        scriptMs: ms(spent.script),
        styleMs: ms(spent.style),
        layoutMs: ms(spent.layout),
        /** Main-thread milliseconds burned per streamed delta. */
        scriptMsPerDelta: Math.round((ms(spent.script) / (CHUNKS * 2)) * 100) / 100,
      }),
      '<<<END',
    )

    // A measurement with no frames is no measurement: a window that never paints
    // would show zero mutations for the wrong reason.
    expect(probe.frames).toBeGreaterThan(10)

    // THE load-bearing assertion, and deliberately a *ratio*: commits per frame.
    // A slow machine lowers frames and commits together, so the ratio holds and
    // the test does not flake; removing the coalescer raises it, and no machine
    // speed excuses that. Measured over 5 runs (deltas 8 ms apart ≈ 125 tokens/s,
    // the local endpoint's decode speed):
    //
    //   arm             mutations / frame      mutations     script
    //   coalesced       0.866 – 0.909          209 – 219     632 – 738 ms
    //   per-delta       1.642 – 1.665          394 – 398    1045 – 1068 ms
    //
    // 1.2 sits between: a commit may touch a few text nodes (hence >1) but must
    // never cost a frame of its own.
    expect(probe.mutations).toBeLessThanOrEqual(Math.ceil(probe.frames * 1.2))
    // Absolute numbers are recorded and guarded loosely — a loaded machine
    // inflates them together, so these catch a pathological regression (a commit
    // that got much more expensive), not the removal of the coalescer. That is
    // the ratio's job above.
    expect(probe.mutations).toBeLessThan(CHUNKS * 2 * 0.6)
    expect(ms(spent.script)).toBeLessThan(2_000)

    // 2.4: the reasoning the IPC chain used to drop, now on screen…
    const disclosure = sheets.locator('details.ai-reasoning')
    await expect(disclosure.locator('.ai-reasoning-body')).toContainText(REASONING.trim())
    // …and folded away now that the answer exists, so the transcript stays
    // readable while still being there to re-open.
    expect(await disclosure.evaluate((el) => (el as HTMLDetailsElement).open)).toBe(false)
    await sheets.screenshot({ path: '/tmp/streaming.png' })
  } finally {
    await launched.app.close().catch(() => {})
    server.close()
  }
})
