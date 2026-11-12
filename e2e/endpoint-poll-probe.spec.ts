import { test } from '@playwright/test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { buildGridWorkbook, launchShell, waitForPageWithUrl } from './helpers'

/**
 * Probe, not a regression test. Report: with GenOffice open and nothing
 * happening on screen, the model gateway keeps getting hammered by
 * `GET /metrics?requests=all`. The app never sends that path — no version in
 * the repo's history contains it — so this probe measures what the app
 * ACTUALLY puts on the wire against its configured endpoint, through a
 * recording proxy that forwards to the real gateway (a real upstream, so real
 * failure modes are exercised, unlike a stub that 404s).
 *
 * Phases mirror the report: cold idle, then idle with a workbook open, then a
 * real AI turn, then idle again. The artifact is the request histogram with
 * User-Agent per route: /tmp/endpoint-poll-probe.json.
 *
 * Point GENOFFICE_PROBE_UPSTREAM at `host:port` (default 127.0.0.1:8899, a
 * recorder forwarding to the gateway) and run with DISPLAY set.
 */

const UPSTREAM = process.env.GENOFFICE_PROBE_UPSTREAM ?? '127.0.0.1:8899'
const SRC = process.env.GENOFFICE_E2E_XLSX ?? ''
const IDLE_MS = Number(process.env.GENOFFICE_PROBE_IDLE_MS ?? 20_000)

interface Hit {
  atMs: number
  dir: string
  method?: string
  url?: string
  ua?: string
  status?: number
  bodyBytes?: number
  note?: string
}

async function readHits(path: string): Promise<Hit[]> {
  if (!existsSync(path)) return []
  return (await readFile(path, 'utf8'))
    .split('\n')
    .filter(Boolean)
    .map((line) => JSON.parse(line) as Hit)
    .filter((h) => h.dir === 'req')
}

function histogram(hits: Hit[]): Array<{ route: string; count: number; userAgents: string[] }> {
  const map = new Map<string, { count: number; userAgents: string[] }>()
  for (const hit of hits) {
    const key = `${hit.method} ${hit.url}`
    const entry = map.get(key) ?? { count: 0, userAgents: [] }
    entry.count += 1
    if (entry.userAgents.length < 3 && !entry.userAgents.includes(hit.ua ?? ''))
      entry.userAgents.push(hit.ua ?? '')
    map.set(key, entry)
  }
  return [...map.entries()].map(([route, e]) => ({ route, ...e }))
}

test('what the app really sends to its AI endpoint, per phase', async () => {
  // A probe, not a regression test: it needs a recording proxy listening at
  // GENOFFICE_PROBE_UPSTREAM. Skip unless one was asked for explicitly, so the
  // shared suite never runs (or reddens on) it by accident.
  test.skip(
    !process.env.GENOFFICE_PROBE_UPSTREAM,
    'set GENOFFICE_PROBE_UPSTREAM=host:port (a recording proxy in front of the gateway)',
  )
  test.setTimeout(300_000)
  const hitsPath = '/tmp/poll-probe/hits.jsonl'

  // Teeth: without the recorder, every phase below reads 0 and the probe would
  // report "the app sends nothing" — a false acquittal. So calibrate against a
  // nonce only this run could produce: a stale hits file from an earlier run
  // (or a restarted proxy, whose atMs clock restarts) cannot satisfy it.
  // The calibration line is also the anchor every phase slices from, by line
  // number — stale lines necessarily precede it.
  const nonce = `${Date.now()}-${Math.random().toString(36).slice(2)}`
  await fetch(`http://${UPSTREAM}/v1/models?probe=${nonce}`).catch(() => {})
  const calibrationIndex = async () =>
    (await readHits(hitsPath)).findIndex((h) => h.url?.includes(nonce))
  const deadline = Date.now() + 5_000
  let anchor = await calibrationIndex()
  while (anchor < 0 && Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, 200))
    anchor = await calibrationIndex()
  }
  if (anchor < 0) {
    throw new Error(
      `recording proxy at ${UPSTREAM} did not log this run's calibration request — ` +
        `start it (proxy-record.cjs) before running this probe, otherwise "0 requests" means nothing`,
    )
  }

  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-poll-probe-'))
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = {
    apiKey: 'probe',
    // A model the gateway actually serves, so phase 3 exercises a real
    // completion (and whatever the provider does after one) rather than a 4xx.
    model: process.env.GENOFFICE_PROBE_MODEL ?? 'qwen3.8-flash-next-unsloth-q5_k_xl',
    baseUrl: `http://${UPSTREAM}/v1`,
  }
  await writeFile(join(userDataDir, 'ai-settings.json'), JSON.stringify(resolveCustomEndpoints(settings)))

  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-poll-book-'))
  const workbook = join(scratch, 'book.xlsx')
  await writeFile(
    workbook,
    SRC
      ? await readFile(SRC)
      : await buildGridWorkbook([
          ['name', 'amount'],
          ['alpha', '10'],
          ['beta', '20'],
          ['gamma', '30'],
        ]),
  )

  // Phase boundaries are line numbers in the recorder's log, not timestamps: a
  // phase with zero hits must report zero (not swallow the next phase's), and
  // atMs is only comparable within one proxy process.
  const boundary = async () => (await readHits(hitsPath)).length
  const sliceFrom = async (from: number) => (await readHits(hitsPath)).slice(from)
  const phases: Array<{ phase: string; hits: Hit[] }> = []
  const report = (phase: string, hits: Hit[]) => {
    phases.push({ phase, hits })
    console.log(`PHASE ${phase}: ${hits.length} requests — ${JSON.stringify(histogram(hits))}`)
  }

  // Phase 1: cold start, no document, nothing clicked.
  let launched = await launchShell({ userDataDir, onboardingSeen: true, videoDir: 'poll-probe' })
  try {
    let from = await boundary()
    await new Promise((r) => setTimeout(r, IDLE_MS))
    report('1-cold-idle-no-document', await sliceFrom(from))

    // Phase 2: a workbook is open, still nothing clicked. A second launch is
    // the honest way to get there — the same path a double-clicked file takes
    // (argv → open) — and the process boundary keeps the phase split clean.
    await launched.app.close().catch(() => {})
    launched = await launchShell({
      userDataDir,
      onboardingSeen: true,
      openFile: workbook,
      videoDir: 'poll-probe',
    })
    const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
    await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    from = await boundary()
    await new Promise((r) => setTimeout(r, IDLE_MS))
    report('2-idle-with-document', await sliceFrom(from))

    // Phase 3: one real AI turn.
    const composer = sheets.getByPlaceholder(/Ask AI to analyze/)
    await composer.click()
    await composer.fill('How many rows are there? Just answer.')
    await sheets.getByRole('button', { name: 'Send', exact: true }).click()
    from = await boundary()
    await sheets.waitForTimeout(25_000)
    report('3-one-ai-turn', await sliceFrom(from))

    // Phase 4: done talking, idle again. The reported scenario.
    from = await boundary()
    await new Promise((r) => setTimeout(r, IDLE_MS))
    report('4-idle-after-turn', await sliceFrom(from))

    const all = phases.flatMap((p) => p.hits)
    await writeFile(
      '/tmp/endpoint-poll-probe.json',
      JSON.stringify(
        {
          upstream: UPSTREAM,
          idleMs: IDLE_MS,
          totalRequests: all.length,
          phases: phases.map((p) => ({ phase: p.phase, total: p.hits.length, histogram: histogram(p.hits) })),
          all,
        },
        null,
        2,
      ),
    )
  } finally {
    await launched.app.close().catch(() => {})
  }
})
