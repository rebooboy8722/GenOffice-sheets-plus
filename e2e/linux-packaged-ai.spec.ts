import { expect, test, type Page } from '@playwright/test'
import { execSync } from 'node:child_process'
import { existsSync, mkdtempSync } from 'node:fs'
import { createServer, request as httpRequest } from 'node:http'
import type { AddressInfo } from 'node:net'
import { mkdir, readFile, readdir, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { buildGridWorkbook, launchShell, screenshotPath, waitForPageWithUrl } from './helpers'

/**
 * The AI chain in the **packaged** Linux build, against the real local endpoint
 * (127.0.0.1) — the same one the dev-mode real-write spec uses, proving the
 * packaged artifact reaches it too: settings persist into the redirected profile,
 * a real model round trip writes the real grid, the save lands in the file, and
 * an image attachment rides along as vision input.
 *
 * Skips when the endpoint is unreachable, so CI without the LAN is unaffected.
 *
 * **What this proves and what it does not.** Evidence is the AI trace (one JSONL
 * line per provider turn) and the grid/file — never model prose. The turns are
 * stochastic: per the method rules this spec must not be cited as the teeth for
 * any single product change (the deterministic gate tests are
 * `linux-freshness-packaged.spec.ts`). What it earns is the end-to-end claim:
 * *this artifact, on this OS, with a real model, does the work.*
 *
 *   GENOFFICE_LINUX_SMOKE=1 DISPLAY=:77 npx playwright test \
 *     --config e2e/playwright.config.ts e2e/linux-packaged-ai.spec.ts
 */

const ENABLED = process.platform === 'linux' && process.env.GENOFFICE_LINUX_SMOKE === '1'
const SHELL_DIR = resolve(__dirname, '../apps/shell')
const RELEASE = process.env.GENOFFICE_LINUX_DIST ?? join(SHELL_DIR, 'release')
const EXE = join(RELEASE, 'linux-unpacked/genoffice')
const BASE_URL = process.env.GENOFFICE_E2E_AI_URL ?? 'http://127.0.0.1:8080/v1'
const MODEL = process.env.GENOFFICE_E2E_AI_MODEL ?? 'Qwen3.6-27B-UD-Q5_K_XL.gguf'

const freshHome = () => mkdtempSync(join(tmpdir(), 'genoffice-linux-ai-'))

/** The 4×4 literal block from the dev-mode real-write spec (same reasoning: a
 * block of prose is the only task shape that forces set_range over formulas). */
const WRITE_TASK =
  'Write this 4x4 block of text into B7:E10 exactly as written here, one item per cell, ' +
  'row by row: north, south, east, west / kilo, lima, oscar, papa ' +
  '/ red, green, blue, amber / swift, brisk, calm, keen. ' +
  'These are plain text labels, not formulas. Do not change any other cell.'
const BLOCK = [
  ['north', 'south', 'east', 'west'],
  ['kilo', 'lima', 'oscar', 'papa'],
  ['red', 'green', 'blue', 'amber'],
  ['swift', 'brisk', 'calm', 'keen'],
]

const DATA: Array<Array<string | number>> = [
  ['name', 'qty'],
  ['alpha', 1],
  ['beta', 2],
  ['gamma', 3],
]

/** One cell as text through the debug facade. Empty when the facade is not up
 * yet — the debug hook lands after the renderer boots, and callers poll. */
function cellText(page: Page, row: number, col: number): Promise<string> {
  return page
    .evaluate(
      ({ row, col }) => {
        const win = window as unknown as {
          __genofficeDebug?: { univerAPI: { getActiveWorkbook(): any } }
        }
        const v = win.__genofficeDebug?.univerAPI
          ?.getActiveWorkbook()
          ?.getActiveSheet()
          ?.getRange(row, col, 1, 1)
          .getValue()
        return v === null || v === undefined ? '' : String(v)
      },
      { row, col },
    )
    .catch(() => '')
}

interface TraceLine {
  messages: Array<{
    role: string
    results?: Array<{ toolName?: string; output?: unknown; isError?: boolean }>
  }>
  response: { toolCalls: Array<{ name: string; input?: unknown }>; error?: string }
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
        /* partially written line; the poll retries */
      }
    }
  }
  return lines
}

/** Run-over detection from the trace + the Stop button (real-write spec's rule):
 * the last turn asked for no tools and the trace has been quiet for 3 s. */
async function waitForRunEnd(
  dir: string,
  page: Page,
  turnOffset: number,
  timeoutMs = 300_000,
): Promise<TraceLine[]> {
  const deadline = Date.now() + timeoutMs
  let lastChange = Date.now()
  let lastCount = -1
  for (;;) {
    const lines = await readTrace(dir)
    if (lines.length !== lastCount) {
      lastCount = lines.length
      lastChange = Date.now()
    }
    const idle = (await page.getByRole('button', { name: 'Stop', exact: true }).count()) === 0
    const last = lines.at(-1)
    if (
      lines.length > turnOffset &&
      last &&
      last.response.toolCalls.length === 0 &&
      idle &&
      Date.now() - lastChange > 3_000
    ) {
      return lines
    }
    if (Date.now() > deadline) return lines
    await new Promise((r) => setTimeout(r, 1_000))
  }
}

const flattenOutputs = (lines: TraceLine[]): string[] =>
  lines.flatMap((l) => (l.messages ?? []).flatMap((m) => (m.results ?? []).map((r) => String(r.output ?? ''))))

/** Parts are deflate-compressed, so search the unzipped sheet, not the zip. */
const savedSheet = (path: string): string => {
  try {
    return execSync(`unzip -p "${path}" xl/worksheets/sheet1.xml`).toString()
  } catch {
    return ''
  }
}

async function seedRealEndpoint(homeDir: string): Promise<string> {
  const profileDir = join(homeDir, '.config', 'GenOffice')
  await mkdir(profileDir, { recursive: true })
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = { apiKey: 'local', model: MODEL, baseUrl: BASE_URL }
  await writeFile(join(profileDir, 'ai-settings.json'), JSON.stringify(resolveCustomEndpoints(settings)))
  return profileDir
}

async function askAi(page: Page, text: string): Promise<void> {
  const composer = page.getByPlaceholder(/Ask AI to analyze/)
  await composer.click()
  await composer.fill(text)
  await page.getByRole('button', { name: 'Send', exact: true }).click()
}

test.describe('linux packaged AI chain (real endpoint)', () => {
  test.skip(!ENABLED, 'set GENOFFICE_LINUX_SMOKE=1 (linux only)')
  test.skip(!existsSync(EXE), `no packaged build at ${EXE} — run npm run dist:linux`)

  // Probed once per file; each test skips on the flag. (A top-level `await`
  // here would be a syntax error — describe callbacks are synchronous.)
  let endpointOk = false
  test.beforeAll(async () => {
    endpointOk = await fetch(`${BASE_URL.replace(/\/v1$/, '')}/v1/models`, {
      signal: AbortSignal.timeout(6_000),
    })
      .then((r) => r.ok)
      .catch(() => false)
  })

  test('C1 the packaged settings pane shows the seeded 210 endpoint', async () => {
    test.skip(!endpointOk, `model endpoint unreachable at ${BASE_URL}`)
    const homeDir = freshHome()
    await seedRealEndpoint(homeDir)
    const launched = await launchShell({
      packaged: { executable: EXE, homeDir },
      onboardingSeen: true,
      settings: { starPrompt: { resolved: true } },
      videoDir: 'linux-ai',
    })
    try {
      const shell = launched.page
      await shell.getByRole('button', { name: 'Settings', exact: true }).click()
      await shell.locator('.set-nav-item').filter({ hasText: 'AI Model' }).click()
      await shell.getByRole('button', { name: 'Provider', exact: true }).click()
      await shell.getByRole('option', { name: 'Custom', exact: true }).click()
      // The value the pane renders came from the fake home's ai-settings.json —
      // so the packaged app really reads the redirected profile, and a user on
      // this build would see their LAN endpoint, not a default.
      await expect(shell.locator('#set-ai-base-url')).toHaveValue(BASE_URL, { timeout: 15_000 })
      await shell.screenshot({ path: screenshotPath('linux-C1-settings-endpoint') })
    } finally {
      await launched.app.close().catch(() => {})
    }
  })

  test('C2–C4 a real model reads the sheet, writes a 4×4 block with zero rejections, and it lands in the file', async () => {
    test.skip(!endpointOk, `model endpoint unreachable at ${BASE_URL}`)
    test.setTimeout(600_000)
    const homeDir = freshHome()
    await seedRealEndpoint(homeDir)
    const workbook = join(homeDir, 'book.xlsx')
    await writeFile(workbook, await buildGridWorkbook(DATA))
    const traceDir = join(homeDir, 'trace')
    await mkdir(traceDir, { recursive: true })

    const launched = await launchShell({
      packaged: { executable: EXE, homeDir },
      onboardingSeen: true,
      openFile: workbook,
      settings: { starPrompt: { resolved: true } },
      videoDir: 'linux-ai',
      env: {
        GENOFFICE_DEBUG_HOOKS: '1',
        GENOFFICE_AI_TRACE_DIR: traceDir,
        no_proxy: '127.0.0.1,localhost,127.0.0.1',
        NO_PROXY: '127.0.0.1,localhost,127.0.0.1',
      },
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
      // Real content, not the Sheet1 tab label (B5's lesson).
      await expect
        .poll(() => cellText(sheets, 0, 0), { timeout: 40_000, intervals: [500] })
        .toBe('name')

      // ── C2: a read question produces real tool-call turns ─────────────────
      await askAi(sheets, 'What text is in cell A1? Answer with just the text.')
      const readRun = await waitForRunEnd(traceDir, sheets, 0)
      const readTools = readRun.flatMap((l) => l.response.toolCalls.map((c) => c.name))
      console.log('C2 tools:', JSON.stringify(readTools))
      // Never the model's wording — only that the packaged renderer really drove
      // a tool round trip against the real endpoint.
      expect(readTools.length).toBeGreaterThan(0)

      // ── C3: the 4×4 write; a batch rejection would zero the whole run ─────
      const offset = readRun.length
      await askAi(sheets, WRITE_TASK)
      const writeRun = await waitForRunEnd(traceDir, sheets, offset)
      const writeLines = writeRun.slice(offset)
      const outputs = flattenOutputs(writeLines)
      const rejected = outputs.filter((o) => /Rejected — none of/i.test(o)).length
      console.log(
        'C3 turns:',
        writeLines.length,
        'ops:',
        JSON.stringify(
          writeLines.flatMap((l) =>
            l.response.toolCalls.flatMap((c) => {
              const input = c.input as { operations?: Array<{ op?: string }> } | undefined
              return (input?.operations ?? []).map((o) => o.op ?? '?')
            }),
          ),
        ),
      )
      // The teeth the user reported: an all-or-nothing batch with invented field
      // names burned a whole round trip in the shipped build.
      expect(rejected).toBe(0)
      for (const [r, row] of BLOCK.entries()) {
        for (const [c, value] of row.entries()) {
          expect(await cellText(sheets, 6 + r, 1 + c), `B7:E10 cell r${r} c${c}`).toBe(value)
        }
      }
      await sheets.screenshot({ path: screenshotPath('linux-C3-after-write') })

      // ── C4: the packaged save path writes the AI's work to disk ────────────
      await launched.app.evaluate(({ webContents }) => {
        webContents
          .getAllWebContents()
          .find((w) => w.getURL().includes('://sheets/'))
          ?.send('menu:action', 'save')
      })
      await expect
        .poll(() => savedSheet(workbook), { timeout: 30_000, intervals: [1_000] })
        .toContain('oscar')
    } finally {
      await launched.app.close().catch(() => {})
    }
  })

  test('C5 Home Ask AI carries an attached PNG to the endpoint as vision input', async () => {
    test.skip(!endpointOk, `model endpoint unreachable at ${BASE_URL}`)
    test.setTimeout(240_000)
    const homeDir = freshHome()
    const photoPath = join(homeDir, 'photo.png')
    // 1×1 red PNG — the payload shape is what's under test, not its content.
    await writeFile(
      photoPath,
      Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
        'base64',
      ),
    )

    // A recording forward-proxy in front of the real endpoint: the claim is
    // about the *outgoing request bytes* (the trace stores the provider-neutral
    // message list, not the OpenAI wire form), and the answer must still come
    // from the real model — so the proxy forwards and pipes the response back.
    const captured: string[] = []
    // The upstream status matters as much as the capture: a request the proxy
    // saw but 210 rejected would still satisfy "image_url on the wire" while the
    // image never reached a model. Record it and assert 200 below.
    let upstreamStatus = 0
    const target = new URL(BASE_URL)
    const proxy = createServer((req, res) => {
      const chunks: Buffer[] = []
      req.on('data', (c) => chunks.push(c as Buffer))
      req.on('end', () => {
        const body = Buffer.concat(chunks)
        const isChat = !!req.url?.includes('/chat/completions')
        if (isChat) captured.push(body.toString('utf8'))
        const upstream = httpRequest(
          {
            hostname: target.hostname,
            port: target.port,
            path: req.url,
            method: req.method,
            headers: { ...req.headers, host: `${target.hostname}:${target.port}` },
          },
          (up) => {
            if (isChat) upstreamStatus = up.statusCode ?? 0
            res.writeHead(up.statusCode ?? 502, up.headers)
            up.pipe(res)
          },
        )
        upstream.on('error', (e) => res.writeHead(502).end(String(e)))
        upstream.end(body)
      })
    })
    await new Promise<void>((r) => proxy.listen(0, '127.0.0.1', () => r()))
    const port = (proxy.address() as AddressInfo).port

    // Seed through the proxy instead of 210 directly — everything else identical.
    const profileDir = join(homeDir, '.config', 'GenOffice')
    await mkdir(profileDir, { recursive: true })
    const settings = defaultAiSettings()
    settings.provider = 'custom'
    settings.providers.custom = { apiKey: 'local', model: MODEL, baseUrl: `http://127.0.0.1:${port}/v1` }
    await writeFile(join(profileDir, 'ai-settings.json'), JSON.stringify(resolveCustomEndpoints(settings)))

    const launched = await launchShell({
      packaged: { executable: EXE, homeDir },
      onboardingSeen: true,
      settings: { starPrompt: { resolved: true } },
      videoDir: 'linux-ai',
      env: { no_proxy: '127.0.0.1,localhost', NO_PROXY: '127.0.0.1,localhost' },
    })
    try {
      const shell = launched.page
      await shell.getByRole('button', { name: 'Ask AI', exact: true }).click()
      const panel = shell.locator('.kb-ask')
      await expect(panel).toBeVisible({ timeout: 20_000 })
      // Drop via a real <input type=file> — the only way to mint File objects
      // the preload can resolve to paths (home-ask-attachments.spec.ts precedent).
      await shell.evaluate(() => {
        const input = document.createElement('input')
        input.type = 'file'
        input.multiple = true
        input.id = 'e2e-files'
        input.style.display = 'none'
        document.body.appendChild(input)
      })
      await shell.locator('#e2e-files').setInputFiles(photoPath)
      await panel.evaluate((el) => {
        const input = document.getElementById('e2e-files') as HTMLInputElement
        const dt = new DataTransfer()
        for (const f of Array.from(input.files ?? [])) dt.items.add(f)
        for (const type of ['dragenter', 'dragover', 'drop'] as const) {
          el.dispatchEvent(new DragEvent(type, { bubbles: true, cancelable: true, dataTransfer: dt }))
        }
      })
      await expect(panel.locator('.kb-ask-attachments .kb-ask-attachment')).toHaveCount(1)
      await shell.screenshot({ path: screenshotPath('linux-C5-ask-attachment') })
      await panel.locator('textarea').fill('What color is this image?')
      await panel.locator('textarea').press('Enter')

      await expect
        .poll(() => captured.length, { timeout: 60_000, intervals: [1_000] })
        .toBeGreaterThan(0)
      const first = captured[0]
      expect(first).toContain('image_url')
      expect(first).toContain('data:image/png;base64,')
      // The endpoint must have *accepted* it — a 4xx/5xx would still put
      // image_url on the wire while the image reached no model.
      await expect.poll(() => upstreamStatus, { timeout: 120_000, intervals: [1_000] }).toBe(200)
      // And a real assistant turn landed. Two false-greens were caught here by
      // the screenshot: (a) `not.toBeEmpty()` passes on the streaming placeholder
      // "The AI is answering…", and (b) polling only for "not the placeholder"
      // passes on the empty gap between the placeholder clearing and the first
      // text arriving. One poll that treats both as "not yet": real text only.
      const answer = panel.locator('.kb-ask-msg.kb-ask-assistant').last()
      await expect
        .poll(
          async () => {
            const text = (await answer.innerText().catch(() => '')).trim()
            return /answering/i.test(text) ? '' : text
          },
          { timeout: 120_000, intervals: [1_000] },
        )
        .not.toBe('')
      await shell.screenshot({ path: screenshotPath('linux-C5-ask-answer') })
    } finally {
      await launched.app.close().catch(() => {})
      proxy.close()
    }
  })
})
