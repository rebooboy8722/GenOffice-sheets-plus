import { expect, test } from '@playwright/test'
import { createServer, type Server } from 'node:http'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { AddressInfo } from 'node:net'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { buildGridWorkbook, launchShell, waitForPageWithUrl } from './helpers'

/**
 * The user's report: an AI task was running, then ALL file tabs were closed
 * and the app sat on Home — and the model kept getting hammered (big prefill +
 * decode chains, stopping only when the app itself was killed). 148's report
 * blamed a phantom /metrics poller; the real question this spec answers is
 * narrow and decisive:
 *
 *   does an in-flight agent run keep issuing requests after its tab is closed?
 *
 * The stub answers EVERY turn with another (harmless, read-only) tool call, so
 * the run never ends on its own — a runaway chain by construction, the shape
 * the user saw. Once the chain is demonstrably in flight (>= 3 requests), the
 * test closes the file tab through the real UI and counts further POSTs.
 *
 * Expected (and asserted): the loop lives in the tab's renderer, closing the
 * tab destroys it (`destroyView → webContents.close()`), so at most the ONE
 * in-flight completion may finish — no NEW requests. If this test fails, the
 * user's "关了文件页还在转" is reproduced deterministically and the fix is to
 * abort the run on tab teardown.
 */

function startEndlessStub(port0: { count: number }): Promise<{ server: Server; stop: () => void }> {
  const server = createServer((req, res) => {
    let body = ''
    req.on('data', (c: Buffer) => (body += c.toString()))
    req.on('end', () => {
      if (req.method === 'POST' && req.url?.endsWith('/chat/completions')) {
        port0.count += 1
        // A valid read-only tool call every turn: the app executes it and asks
        // again, forever. get_workbook_context takes no arguments.
        const message = {
          role: 'assistant',
          content: null,
          tool_calls: [{ id: `call_${port0.count}`, type: 'function', function: { name: 'get_workbook_context', arguments: '{}' } }],
        }
        res.writeHead(200, { 'content-type': 'application/json' })
        res.end(JSON.stringify({ choices: [{ message, finish_reason: 'tool_calls' }] }))
        return
      }
      void body
      res.writeHead(404, { 'content-type': 'application/json' })
      res.end(JSON.stringify({ error: 'not found' }))
    })
  })
  return new Promise((resolve) => {
    server.listen(0, '127.0.0.1', () => resolve({ server, stop: () => server.close() }))
  })
}

test('a closed file tab must stop the agent run it was hosting', async () => {
  const stub = { count: 0 }
  const { server, stop } = await startEndlessStub(stub)
  const port = (server.address() as AddressInfo).port

  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-runclose-'))
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = { apiKey: 'stub', model: 'stub-model', baseUrl: `http://127.0.0.1:${port}/v1` }
  await writeFile(join(userDataDir, 'ai-settings.json'), JSON.stringify(resolveCustomEndpoints(settings)))

  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-runclose-book-'))
  const workbook = join(scratch, 'book.xlsx')
  await writeFile(
    workbook,
    await buildGridWorkbook([
      ['name', 'amount'],
      ['alpha', '10'],
      ['beta', '20'],
    ]),
  )

  const launched = await launchShell({ userDataDir, onboardingSeen: true, openFile: workbook, videoDir: 'run-close' })
  try {
    const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
    await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, { timeout: 30_000 })

    const composer = sheets.getByPlaceholder(/Ask AI to analyze/)
    await composer.click()
    await composer.fill('take a look at this workbook')
    await sheets.getByRole('button', { name: 'Send', exact: true }).click()

    // Wait until the runaway chain is demonstrably in flight.
    await expect
      .poll(() => stub.count, { timeout: 60_000, message: 'runaway chain never started' })
      .toBeGreaterThanOrEqual(3)

    // Close the file tab through the real UI (shell window renders the tab bar).
    const shell = launched.page
    const dump = await shell
      .evaluate(() => {
        const bar = document.querySelector('.tab-bar')
        return {
          hasBar: !!bar,
          closeButtons: document.querySelectorAll('.tab-close').length,
          barHtml: bar ? bar.outerHTML.slice(0, 400) : document.body.className,
        }
      })
      .catch((e: unknown) => ({ hasBar: false, closeButtons: -1, barHtml: String(e).slice(0, 200) }))
    console.log(`TABBAR: ${JSON.stringify(dump)}`)
    await shell.locator('.tab-close').first().click()
    // Home is back: the sheets view must be gone.
    await expect.poll(() => sheets.isClosed()).toBe(true)

    // The in-flight completion may finish; a NEW request must not start.
    const atClose = stub.count
    await new Promise((r) => setTimeout(r, 20_000))
    const after = stub.count
    console.log(
      `RUN-AFTER-CLOSE: ${atClose} requests before close, ${after} after 20 s ` +
        `(new: ${after - atClose})`,
    )
    expect(after - atClose, 'closed tab kept issuing new model requests').toBeLessThanOrEqual(1)
  } finally {
    await launched.app.close().catch(() => {})
    stop()
  }
})
