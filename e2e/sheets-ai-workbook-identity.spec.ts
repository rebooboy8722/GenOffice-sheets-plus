import { expect, test, type ElectronApplication } from '@playwright/test'
import { createServer, type Server } from 'node:http'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { execSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { AddressInfo } from 'node:net'
import type { Page } from '@playwright/test'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { buildGridWorkbook, gridCellCenter, launchShell, waitForPageWithUrl } from './helpers'

/**
 * When the model may be told "a different workbook is open" — and when it may not.
 *
 * The freshness baseline notes a workbook switch, because the conversation's
 * earlier ranges then belong to a different file — a drift counter cannot see that
 * (nothing was edited, the workbook was replaced). The note is keyed on the file's
 * on-disk path, and the tempting alternative is wrong here: saving swaps the
 * sidecar session and rebuilds the Univer unit for the *same* path, so keying on
 * the unit id or session id would tell the model "a DIFFERENT WORKBOOK IS NOW
 * OPEN, do NOT reuse those addresses" after every save, on a file that never
 * changed.
 *
 * Three ways the identity moves, three tests, all reading the request body the
 * stub actually received:
 *  - **Save** keeps the path and changes the session/unit → the note must stay
 *    silent (the unit-id trap).
 *  - **Save As** changes the path and is the same document → the note must stay
 *    silent too, which needs the reopen's own "session swap over the same
 *    document" signal (`openLazyWorkbook({continueChat})` → `realignWorkbookIdentity`),
 *    not just a path comparison.
 *  - **File > Open** of another file → the note must be *there*, reported by the
 *    loader (`noteWorkbookDocumentReplaced`), while the loop still carries the
 *    previous file's turns.
 *
 * The two save tests verify the save landed on disk first — a save that never
 * happened would skip the reopen and let the assertion pass vacuously.
 */

interface ChatRequest {
  messages: Array<{ role: string; content: unknown }>
}

function startTextStub(requests: ChatRequest[]): Promise<Server> {
  const server = createServer(async (req, res) => {
    if (req.method !== 'POST') {
      res.writeHead(404).end()
      return
    }
    const chunks: Buffer[] = []
    for await (const chunk of req) chunks.push(chunk as Buffer)
    requests.push(JSON.parse(Buffer.concat(chunks).toString('utf8')) as ChatRequest)
    res.writeHead(200, { 'content-type': 'text/event-stream' })
    res.write(
      `data: ${JSON.stringify({
        choices: [{ index: 0, delta: { role: 'assistant', content: 'Noted.' } }],
      })}\n\n`,
    )
    res.write(
      `data: ${JSON.stringify({ choices: [{ index: 0, delta: {}, finish_reason: 'stop' }] })}\n\n`,
    )
    res.write('data: [DONE]\n\n')
    res.end()
  })
  return new Promise((r) => server.listen(0, '127.0.0.1', () => r(server)))
}

async function ask(page: Page, text: string): Promise<void> {
  const composer = page.getByPlaceholder(/Ask AI to analyze/)
  await composer.click()
  await composer.fill(text)
  await page.getByRole('button', { name: 'Send', exact: true }).click()
  await expect(page.getByText('Noted.', { exact: false }).last()).toBeVisible({ timeout: 30_000 })
}

/** File > Save / Save As, routed to the sheets view the way the app menu does it. */
async function saveThroughMenu(
  app: ElectronApplication,
  mode: 'save' | 'save-as',
  /** Save As answers a native picker; headless, the main process is told the answer. */
  saveTo?: string,
): Promise<void> {
  await app.evaluate(
    ({ webContents, dialog }, args) => {
      if (args.saveTo) {
        dialog.showSaveDialog = async () => ({ canceled: false, filePath: args.saveTo })
      }
      const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('://sheets/'))
      wc?.send('menu:action', args.mode)
    },
    { mode, ...(saveTo ? { saveTo } : {}) },
  )
}

// Parts are deflate-compressed, so the raw zip is not searchable — unzip the two
// parts a string cell can land in (inline vs shared strings).
const savedCells = (path: string): string =>
  ['xl/worksheets/sheet1.xml', 'xl/sharedStrings.xml']
    .map((part) => {
      try {
        return execSync(`unzip -p "${path}" ${part}`).toString()
      } catch {
        return ''
      }
    })
    .join('\n')

/**
 * Boots the app with the stub endpoint; caller closes both. `workbook: null`
 * starts on the home screen and quick-creates a **blank** workbook — one with no
 * path at all, which is the state a path comparison cannot see out of.
 */
async function launchOn(workbook: string | null, videoDir: string, documentsDir?: string) {
  const requests: ChatRequest[] = []
  const server = await startTextStub(requests)
  const port = (server.address() as AddressInfo).port
  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-e2e-'))
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = {
    apiKey: 'stub',
    model: 'stub-model',
    baseUrl: `http://127.0.0.1:${port}/v1`,
  }
  await writeFile(
    join(userDataDir, 'ai-settings.json'),
    JSON.stringify(resolveCustomEndpoints(settings)),
  )
  const launched = await launchShell({
    userDataDir,
    onboardingSeen: true,
    ...(workbook ? { openFile: workbook } : {}),
    videoDir,
    // The File > Open test waits on the live grid; cell text is canvas-drawn, so
    // the only way to read it back is window.__genofficeDebug.univerAPI.
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  if (!workbook) {
    // Home screen: the AI Sheets quick-card is what creates the blank workbook.
    // Keep the auto-created workbook out of the real ~/Documents/GenOffice.
    await launched.app.evaluate(({ app: electronApp }, dir) => {
      electronApp.setPath('documents', dir)
    }, documentsDir ?? (await mkdtemp(join(tmpdir(), 'genoffice-docs-'))))
    await launched.page.locator('.quick-card').nth(1).click()
  }
  const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
  await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
    timeout: 30_000,
  })
  await sheets.waitForTimeout(2_000)
  const close = async (): Promise<void> => {
    await launched.app.close().catch(() => {})
    server.close()
  }
  return { launched, sheets, requests, close }
}

/** Types into B2 (row 2) so the save has something to carry — see the note below. */
async function editACell(sheets: Page, text: string): Promise<void> {
  const cell = await gridCellCenter(sheets, 1, 0)
  await sheets.mouse.click(cell.x, cell.y)
  await sheets.keyboard.type(text, { delay: 40 })
  await sheets.keyboard.press('Enter')
  await sheets.waitForTimeout(500)
}

test('saving the same workbook does not tell the model the workbook changed', async () => {
  test.setTimeout(180_000)
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-save-note-'))
  const workbook = join(scratch, 'book.xlsx')
  await writeFile(workbook, await buildGridWorkbook([['name'], ['alpha'], ['beta']]))
  const { launched, sheets, requests, close } = await launchOn(workbook, 'save-note')
  try {
    await ask(sheets, 'Remember this workbook.')
    expect(requests).toHaveLength(1)
    expect(JSON.stringify(requests[0])).not.toContain('DIFFERENT WORKBOOK')

    // An edit first: with nothing pending, Save has no session to swap, and the
    // reopen — the event under test — would never happen.
    await editACell(sheets, 'saved-value')

    // Split the hypothesis: does the note fire because of the save, or because
    // the identity changes on its own as the lazy session settles? A plain
    // follow-up turn is the control for that, and worth keeping either way.
    await ask(sheets, 'Still the same file?')
    expect(requests).toHaveLength(2)
    expect(JSON.stringify(requests[1])).not.toContain('DIFFERENT WORKBOOK')

    await saveThroughMenu(launched.app, 'save')
    await expect(() => {
      expect(savedCells(workbook)).toContain('saved-value')
    }).toPass({ timeout: 20_000 })
    await sheets.waitForTimeout(2_000)

    await ask(sheets, 'And now?')
    expect(requests).toHaveLength(3)
    // The regression guard: key the switch note on the Univer unit id or the
    // sidecar session id and this line fails, on a file that never changed.
    expect(JSON.stringify(requests[2])).not.toContain('DIFFERENT WORKBOOK')
  } finally {
    await close()
  }
})

/**
 * The other direction, in the real app: opening a *different* file in this window
 * must tell the model, because the conversation's ranges belong to the file that
 * just closed. Both signals could catch this one (file → file moves the path);
 * the test below is what isolates the loader's signal.
 */
test('opening another workbook tells the model its earlier ranges are stale', async () => {
  test.setTimeout(180_000)
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-open-note-'))
  const first = join(scratch, 'book-a.xlsx')
  const second = join(scratch, 'book-b.xlsx')
  await writeFile(first, await buildGridWorkbook([['alpha-only'], ['alpha']]))
  await writeFile(second, await buildGridWorkbook([['beta-only'], ['beta']]))
  const { launched, sheets, requests, close } = await launchOn(first, 'open-note')
  try {
    await ask(sheets, 'Remember this workbook.')
    expect(requests).toHaveLength(1)
    expect(JSON.stringify(requests[0])).not.toContain('DIFFERENT WORKBOOK')

    // File > Open answers a native picker; headless, the main process is told
    // which file was picked.
    await launched.app.evaluate(({ dialog }, path) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [path] })
    }, second)
    await launched.app.evaluate(({ webContents }) => {
      const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('://sheets/'))
      wc?.send('menu:action', 'open')
    })
    // Wait on the live grid, not the toast: cell text is canvas-drawn (never in
    // the DOM) and the status message auto-dismisses. A1 is 'beta-only' only in
    // book-b, so this proves the second workbook really replaced the first.
    await expect
      .poll(
        async () =>
          sheets.evaluate(() => {
            const api = (
              window as unknown as {
                __genofficeDebug: {
                  univerAPI: {
                    getActiveWorkbook(): {
                      getActiveSheet(): {
                        getRange(r: number, c: number): { getValue(): unknown }
                      } | null
                    } | null
                  }
                }
              }
            ).__genofficeDebug.univerAPI
            return api.getActiveWorkbook()?.getActiveSheet()?.getRange(0, 0)?.getValue() ?? null
          }),
        { timeout: 30_000 },
      )
      .toBe('beta-only')
    await sheets.waitForTimeout(2_000)

    await ask(sheets, 'And now?')
    expect(requests).toHaveLength(2)
    // The loop keeps its history across the open (only "New chat" resets it), so
    // this request still carries book-a's turns — which is exactly why the note
    // has to be there with them.
    expect(JSON.stringify(requests[1])).toContain('DIFFERENT WORKBOOK')
  } finally {
    await close()
  }
})

/**
 * File > Open over a **blank** workbook — the case the path comparison cannot see.
 *
 * A quick-created workbook has no path, so the baseline's identity is null; after
 * the open it is book-b's path, and "null → something" is not a change between two
 * known values. Only the loader knows a document was replaced here
 * (`noteWorkbookDocumentReplaced`). Teeth: drop that call and the note goes missing.
 */
test('opening a file over a blank workbook tells the model too', async () => {
  test.setTimeout(180_000)
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-blank-open-'))
  const second = join(scratch, 'book-b.xlsx')
  await writeFile(second, await buildGridWorkbook([['beta-only'], ['beta']]))
  const { launched, sheets, requests, close } = await launchOn(
    null,
    'blank-open-note',
    join(scratch, 'documents'),
  )
  try {
    await ask(sheets, 'Remember this workbook.')
    expect(requests).toHaveLength(1)
    expect(JSON.stringify(requests[0])).not.toContain('DIFFERENT WORKBOOK')

    await launched.app.evaluate(({ dialog }, path) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [path] })
    }, second)
    await launched.app.evaluate(({ webContents }) => {
      const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('://sheets/'))
      wc?.send('menu:action', 'open')
    })
    await expect
      .poll(
        async () =>
          sheets.evaluate(() => {
            const api = (
              window as unknown as {
                __genofficeDebug: {
                  univerAPI: {
                    getActiveWorkbook(): {
                      getActiveSheet(): {
                        getRange(r: number, c: number): { getValue(): unknown }
                      } | null
                    } | null
                  }
                }
              }
            ).__genofficeDebug.univerAPI
            return api.getActiveWorkbook()?.getActiveSheet()?.getRange(0, 0)?.getValue() ?? null
          }),
        { timeout: 30_000 },
      )
      .toBe('beta-only')
    await sheets.waitForTimeout(2_000)

    await ask(sheets, 'And now?')
    expect(requests).toHaveLength(2)
    expect(JSON.stringify(requests[1])).toContain('DIFFERENT WORKBOOK')
  } finally {
    await close()
  }
})

test('Save As to a new path does not tell the model the workbook changed', async () => {
  test.setTimeout(180_000)
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-saveas-note-'))
  const workbook = join(scratch, 'draft.xlsx')
  const renamed = join(scratch, 'final.xlsx')
  await writeFile(workbook, await buildGridWorkbook([['name'], ['alpha'], ['beta']]))
  const { launched, sheets, requests, close } = await launchOn(workbook, 'saveas-note')
  try {
    await ask(sheets, 'Remember this workbook.')
    expect(requests).toHaveLength(1)

    await editACell(sheets, 'renamed-value')
    await saveThroughMenu(launched.app, 'save-as', renamed)
    // The reopen is the event under test, so first prove the Save As really
    // happened: the new file holds the edit, the old one does not.
    await expect(() => {
      expect(savedCells(renamed)).toContain('renamed-value')
    }).toPass({ timeout: 20_000 })
    expect(savedCells(workbook)).not.toContain('renamed-value')
    await sheets.waitForTimeout(2_000)

    await ask(sheets, 'And now?')
    expect(requests).toHaveLength(2)
    // Path-only keying fails here: Save As moved the path while the document
    // stayed the same, and the reopen's continueChat signal is what excuses it.
    expect(JSON.stringify(requests[1])).not.toContain('DIFFERENT WORKBOOK')
  } finally {
    await close()
  }
})
