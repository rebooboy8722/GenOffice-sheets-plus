import { expect, test, type ElectronApplication, type Page } from '@playwright/test'
import { existsSync, mkdtempSync } from 'node:fs'
import { readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { buildGridWorkbook, launchShell, screenshotPath, waitForPageWithUrl } from './helpers'

/**
 * Packaged-Linux smoke: drive the **installed-layout** build (not
 * `apps/shell/out`) and check the user-visible basics, one screenshot per step.
 *
 * Gated so the shared suite is untouched on mac/win/CI:
 *   GENOFFICE_LINUX_SMOKE=1 DISPLAY=:77 npx playwright test \
 *     --config e2e/playwright.config.ts e2e/linux-packaged-smoke.spec.ts
 *
 * Why a separate launcher path is needed at all: `GENOFFICE_USER_DATA` is
 * ignored once `app.isPackaged` (see `e2e/helpers.ts` `packaged` docs), so the
 * run is isolated by a scratch HOME instead. Each test gets its own home ⇒ no
 * profile reuse, and the single-instance lock can never reach a real profile.
 */

const ENABLED = process.platform === 'linux' && process.env.GENOFFICE_LINUX_SMOKE === '1'
const SHELL_DIR = resolve(__dirname, '../apps/shell')
const RELEASE = process.env.GENOFFICE_LINUX_DIST ?? join(SHELL_DIR, 'release')
const EXE = join(RELEASE, 'linux-unpacked/genoffice')
const VERSION = process.env.GENOFFICE_LINUX_VERSION ?? '0.12.29-fork-r3linux'

const freshHome = () => mkdtempSync(join(tmpdir(), 'genoffice-linux-home-'))

/**
 * The shell window is the one exposing `window.aiOffice` (the preload bridge).
 * On a packaged cold boot `firstWindow()` is not reliably the shell — the
 * editor WebContentsView or a splash can come up first — so `launched.page`
 * alone is a coin flip (B4 failed exactly this way in the full run while
 * passing standalone). Poll every window, theme-pipeline.spec.ts's `findShellPage`
 * precedent.
 */
async function findShellPage(app: ElectronApplication, timeoutMs = 30_000): Promise<Page> {
  const deadline = Date.now() + timeoutMs
  for (;;) {
    for (const candidate of app.windows()) {
      const ready = await candidate
        .evaluate(() => typeof (window as unknown as { aiOffice?: unknown }).aiOffice)
        .catch(() => 'detached')
      if (ready === 'object') return candidate
    }
    const remaining = deadline - Date.now()
    if (remaining <= 0) throw new Error('No window exposing window.aiOffice')
    await app.waitForEvent('window', { timeout: Math.min(remaining, 1_000) }).catch(() => {})
  }
}

/**
 * Wait until the workbook's cells are actually in the model. `Sheet1` appears
 * in the DOM the moment the sheet tab renders — *before* the sidecar finishes
 * parsing the file — so a facade read right after that text check can see an
 * empty grid (B5/B6 both failed this way). Poll A1 until it matches the
 * seeded workbook's first cell.
 */
async function waitForCell(
  sheets: Page,
  row: number,
  col: number,
  expected: unknown,
  timeoutMs = 40_000,
): Promise<void> {
  await expect
    .poll(
      () =>
        sheets
          .evaluate(
            ([r, c]) => {
              const debug = (
                window as unknown as {
                  __genofficeDebug: { univerAPI: { getActiveWorkbook(): any } }
                }
              ).__genofficeDebug
              return debug.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getRange(r, c, 1, 1).getValue() ?? null
            },
            [row, col],
          )
          .catch(() => null),
      { timeout: timeoutMs, intervals: [500] },
    )
    .toEqual(expected)
}

/**
 * A workbook whose headers are Chinese and whose amounts are real numbers:
 * it exercises the CJK font subset *inside the grid* (canvas-drawn, so an
 * unresolved font shows as tofu in the screenshot) while giving the tests
 * exact values to read back through the Univer facade.
 */
const SMOKE_ROWS: Array<Array<string | number>> = [
  ['项目', '金额', '数量', '备注'],
  ['外委单甲', 1200, 3, '已结算'],
  ['外委单乙', 4500, 12, '待付款'],
  ['外委单丙', 880, 7, '部分结算'],
]

test.describe('linux packaged smoke', () => {
  test.skip(!ENABLED, 'set GENOFFICE_LINUX_SMOKE=1 (linux only)')
  test.skip(!existsSync(EXE), `no packaged build at ${EXE} — run npm run dist:linux`)

  test('B1 really boots the packaged binary, under a scratch profile', async () => {
    const homeDir = freshHome()
    const launched = await launchShell({
      packaged: { executable: EXE, homeDir },
      onboardingSeen: true,
      videoDir: 'linux-smoke',
    })
    try {
      const facts = await launched.app.evaluate(
        ({ app }, expected) => ({
          packaged: app.isPackaged,
          version: app.getVersion(),
          exec: app.getPath('exe'),
          userData: app.getPath('userData'),
          expected,
        }),
        VERSION,
      )
      console.log('B1 facts:', JSON.stringify(facts))
      // This test is the launcher's own teeth: if the harness ever fell back to
      // `apps/shell/out`, isPackaged would be false and the version 0.12.0, and
      // every other test in this file would be measuring the wrong artifact.
      expect(facts.packaged).toBe(true)
      expect(facts.version).toBe(facts.expected)
      expect(facts.exec).toBe(EXE)
      // isolation actually holds — nothing under the developer's real HOME
      expect(facts.userData.startsWith(homeDir)).toBe(true)
      // and the seeded profile is where the app says it is
      expect(existsSync(join(facts.userData, 'app-settings.json'))).toBe(true)
    } finally {
      await launched.app.close().catch(() => {})
    }
  })

  test('B2 Home renders (hero, quick-create cards, tab strip)', async () => {
    const launched = await launchShell({
      packaged: { executable: EXE, homeDir: freshHome() },
      onboardingSeen: true,
      videoDir: 'linux-smoke',
    })
    try {
      const shell = launched.page
      // Assert the greeting, not the subline: `Home.tsx:45` picks one subline
      // from a pool at random on entry, so any specific one is a coin flip.
      await expect(shell.getByRole('heading', { level: 1 })).toContainText(/Good (morning|afternoon|evening)/, {
        timeout: 30_000,
      })
      // quick-create cards and the tab strip are the shell's own chrome
      await expect(shell.getByRole('region', { name: 'Quick start' })).toBeVisible()
      await expect(shell.getByRole('button', { name: /AI Sheets/ })).toBeVisible()
      await expect(shell.locator('.tab-bar')).toBeVisible()
      // A star prompt can overlay the page; dismiss it so the screenshot is the app
      const later = shell.getByRole('button', { name: /Maybe later/ })
      if (await later.isVisible().catch(() => false)) await later.click()
      await shell.screenshot({ path: screenshotPath('linux-B2-home') })
    } finally {
      await launched.app.close().catch(() => {})
    }
  })

  test('B3 bundled fonts resolve inside the docs editor, including the CJK subset', async () => {
    const launched = await launchShell({
      packaged: { executable: EXE, homeDir: freshHome() },
      onboardingSeen: true,
      videoDir: 'linux-smoke',
    })
    try {
      const shell = launched.page
      // The @font-face rules live in the DOCS renderer (fonts.css), and the
      // ttf/woff2 payloads ship inside the package (resources/modules/docs/
      // renderer/assets/*.ttf). Home has no such rules, so measuring on Home
      // would report an empty font set no matter how well packaging worked.
      await shell.locator('.quick-card').first().click()
      const docs = await waitForPageWithUrl(launched.app, '://docs/')
      await expect(docs.locator('body')).toBeVisible({ timeout: 40_000 })
      await docs.waitForFunction(() => document.fonts.status === 'loaded', null, { timeout: 30_000 })

      const fonts = await docs.evaluate(async () => {
        // force the faces we care about, then report what actually resolved
        await Promise.all([
          document.fonts.load('16px "Carlito GO"'),
          document.fonts.load('16px "Noto Sans CJK SC"'),
        ])
        return {
          carlito: document.fonts.check('16px "Carlito GO"'),
          cjk: document.fonts.check('16px "Noto Sans CJK SC"'),
          loaded: [...document.fonts].filter((f) => f.status === 'loaded').length,
        }
      })
      console.log('B3 fonts:', JSON.stringify(fonts))
      // A Linux box has no Calibri/Cambria and (on a minimal install) no CJK
      // font either — the product's own bundled copies must carry it.
      expect(fonts.carlito).toBe(true)
      expect(fonts.cjk).toBe(true)
      expect(fonts.loaded).toBeGreaterThan(0)

      // CJK must actually occupy width — a missing glyph run renders as tofu
      // boxes (still non-zero width) or nothing at all, so also eyeball this.
      const width = await docs.evaluate(() => {
        const probe = document.createElement('div')
        probe.textContent = '外委单汇总表合计金额'
        probe.style.cssText = 'position:fixed;font:16px "Noto Sans CJK SC";left:-9999px'
        document.body.appendChild(probe)
        const w = probe.getBoundingClientRect().width
        probe.remove()
        return w
      })
      expect(width).toBeGreaterThan(100)

      await docs.screenshot({ path: screenshotPath('linux-B3-docs-fonts') })
    } finally {
      await launched.app.close().catch(() => {})
    }
  })

  test('B4 theme switch reaches shell and editor; document content is not re-authored', async () => {
    const dir = freshHome()
    const workbook = join(dir, 'smoke-book.xlsx')
    await writeFile(workbook, await buildGridWorkbook(SMOKE_ROWS))

    const launched = await launchShell({
      packaged: { executable: EXE, homeDir: dir },
      onboardingSeen: true,
      openFile: workbook,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
      videoDir: 'linux-smoke',
    })
    try {
      // The shell window (preload bridge) — not necessarily firstWindow().
      const shell = await findShellPage(launched.app)
      const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
      // Wait for real content, not the Sheet1 tab label, before the facade read.
      await waitForCell(sheets, 0, 0, '项目')
      const themeOf = (page: Page) =>
        page.evaluate(() => document.documentElement.getAttribute('data-theme'))
      expect(await themeOf(shell)).toBeNull() // default = follow system, no attribute

      // Same control the Settings modal drives (theme-pipeline.spec.ts precedent).
      await shell.evaluate(() => {
        const api = (window as unknown as { aiOffice: { setTheme(v: string): Promise<void> } }).aiOffice
        return api.setTheme('dark')
      })
      await expect.poll(() => themeOf(shell), { timeout: 10_000 }).toBe('dark')
      await expect.poll(() => themeOf(sheets), { timeout: 20_000 }).toBe('dark')
      await shell.screenshot({ path: screenshotPath('linux-B4-shell-dark') })
      await sheets.screenshot({ path: screenshotPath('linux-B4-sheets-dark') })

      // CLAUDE.md rule 4: a dark *page* is a display-time remap only. The
      // authored document data must be untouched by the theme — read the model,
      // not the pixels, so a remap that leaked into the document would fail here.
      const authored = await sheets.evaluate(() => {
        const debug = (
          window as unknown as {
            __genofficeDebug: { univerAPI: { getActiveWorkbook(): any } }
          }
        ).__genofficeDebug
        const sheet = debug.univerAPI.getActiveWorkbook()!.getActiveSheet()!
        return { a1: sheet.getRange(0, 0, 1, 1).getValue(), b2: sheet.getRange(1, 1, 1, 1).getValue() }
      })
      expect(authored).toEqual({ a1: '项目', b2: 1200 })
    } finally {
      await launched.app.close().catch(() => {})
    }
  })

  test('B5 the xlsx sidecar really runs: the file on disk reaches the grid', async () => {
    // In the packaged layout the workbook parser is a *separate binary*
    // (`resources/native/xlsx-sidecar`), and electron-builder exits 0 when that
    // extraResources source is missing. So "the app started" says nothing about
    // "Sheets can open a workbook" — this is the assertion that does.
    const dir = freshHome()
    const workbook = join(dir, 'sidecar.xlsx')
    await writeFile(workbook, await buildGridWorkbook(SMOKE_ROWS))

    const launched = await launchShell({
      packaged: { executable: EXE, homeDir: dir },
      onboardingSeen: true,
      openFile: workbook,
      settings: { starPrompt: { resolved: true } },
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
      videoDir: 'linux-smoke',
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
      // Gate on the real A1, not the Sheet1 tab label — the label renders
      // before the sidecar finishes parsing, which is what made this read nulls.
      await waitForCell(sheets, 0, 0, '项目')
      const grid = await sheets.evaluate(() => {
        const debug = (
          window as unknown as {
            __genofficeDebug: { univerAPI: { getActiveWorkbook(): any } }
          }
        ).__genofficeDebug
        const ws = debug.univerAPI.getActiveWorkbook()!.getActiveSheet()!
        return {
          header: ws.getRange(0, 0, 1, 4).getValues()[0],
          lastRow: ws.getRange(3, 0, 1, 4).getValues()[0],
          rows: ws.getMaxRows(),
          cols: ws.getMaxColumns(),
        }
      })
      console.log('B5 grid:', JSON.stringify(grid))
      expect(grid.header).toEqual(['项目', '金额', '数量', '备注'])
      expect(grid.lastRow).toEqual(['外委单丙', 880, 7, '部分结算'])
      // A dead sidecar yields an empty or demo workbook, not an exception —
      // dimensions are the second, independent signal that real parsing happened.
      expect(grid.rows).toBeGreaterThan(3)
      expect(grid.cols).toBeGreaterThan(3)
      await sheets.screenshot({ path: screenshotPath('linux-B5-sheets-grid') })
    } finally {
      await launched.app.close().catch(() => {})
    }
  })

  test('B6 an edit saves back to its own path and survives a relaunch under a new profile', async () => {
    // Two packaged cold boots in one test — the 90s config default is for one.
    test.setTimeout(240_000)
    const dir = freshHome()
    const workbook = join(dir, 'roundtrip.xlsx')
    await writeFile(workbook, await buildGridWorkbook(SMOKE_ROWS))

    const launched = await launchShell({
      packaged: { executable: EXE, homeDir: dir },
      onboardingSeen: true,
      openFile: workbook,
      settings: { starPrompt: { resolved: true } },
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
      videoDir: 'linux-smoke',
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
      // Load first: setValue on a still-empty grid is overwritten by the async
      // file load, so the doc never goes dirty and Save writes nothing back.
      await waitForCell(sheets, 0, 0, '项目')
      await sheets.evaluate(() => {
        const debug = (
          window as unknown as {
            __genofficeDebug: { univerAPI: { getActiveWorkbook(): any } }
          }
        ).__genofficeDebug
        return debug.univerAPI.getActiveWorkbook()!.getActiveSheet()!.getRange(4, 1, 1, 1).setValue(9999)
      })
      // Same channel the File > Save menu item uses (open-empty-file-save.spec.ts
      // precedent); a workbook opened from a path saves to that path with no dialog.
      await launched.app.evaluate(({ webContents }) => {
        webContents
          .getAllWebContents()
          .find((w) => w.getURL().includes('://sheets/'))
          ?.send('menu:action', 'save')
      })
      const sheetXml = () =>
        readFile(workbook)
          .then(async (bytes) => {
            const { default: JSZip } = await import('jszip')
            const zip = await JSZip.loadAsync(bytes)
            const entry = zip.file('xl/worksheets/sheet1.xml')
            return entry ? entry.async('string') : ''
          })
          .catch(() => '')
      // The bytes on disk, not the in-memory model: proves the packaged save
      // path (sidecar writer + atomic rename) actually wrote this file.
      await expect.poll(sheetXml, { timeout: 30_000, intervals: [1_000] }).toContain('9999')
    } finally {
      await launched.app.close().catch(() => {})
    }

    // Second launch, deliberately a *different* profile: rules out "it only
    // looked saved because the same process still had it".
    const relaunched = await launchShell({
      packaged: { executable: EXE, homeDir: freshHome() },
      onboardingSeen: true,
      openFile: workbook,
      settings: { starPrompt: { resolved: true } },
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
      videoDir: 'linux-smoke',
    })
    try {
      const sheets = await waitForPageWithUrl(relaunched.app, '://sheets/')
      await waitForCell(sheets, 0, 0, '项目')
      const reread = await sheets.evaluate(() => {
        const debug = (
          window as unknown as {
            __genofficeDebug: { univerAPI: { getActiveWorkbook(): any } }
          }
        ).__genofficeDebug
        return debug.univerAPI.getActiveWorkbook()!.getActiveSheet()!.getRange(4, 1, 1, 1).getValue()
      })
      expect(reread).toBe(9999)
    } finally {
      await relaunched.app.close().catch(() => {})
    }
  })

  test('B7 Docs: typed text lands in the .docx and reopens from disk', async () => {
    // Two packaged cold boots in one test — the 90s config default is for one.
    test.setTimeout(240_000)
    const dir = freshHome()
    const doc = join(dir, 'note.docx')
    // A 0-byte .docx opens as a blank document that belongs to its own path
    // (open-empty-file-save.spec.ts), which is what lets Save write in place.
    await writeFile(doc, '')

    const launched = await launchShell({
      packaged: { executable: EXE, homeDir: dir },
      onboardingSeen: true,
      openFile: doc,
      settings: { starPrompt: { resolved: true } },
      videoDir: 'linux-smoke',
    })
    const SENTENCE = 'Linux packaged smoke note'
    try {
      const docs = await waitForPageWithUrl(launched.app, '://docs/')
      await docs.waitForFunction(
        () => Boolean((window as unknown as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
        null,
        { timeout: 40_000 },
      )
      await docs.locator('.doc-page[contenteditable="true"]').waitFor({ timeout: 30_000 })
      await docs.locator('.doc-page').click()
      await docs.keyboard.type(SENTENCE, { delay: 12 })
      await docs.screenshot({ path: screenshotPath('linux-B7-docs-typed') })
      await launched.app.evaluate(({ webContents }) => {
        webContents
          .getAllWebContents()
          .find((w) => w.getURL().includes('://docs/'))
          ?.send('menu:command', 'save')
      })
      const plain = () =>
        readFile(doc)
          .then(async (bytes) => {
            const { default: JSZip } = await import('jszip')
            const zip = await JSZip.loadAsync(bytes)
            const entry = zip.file('word/document.xml')
            // runs can split mid-phrase, so compare text content, not a substring of XML
            return entry ? (await entry.async('string')).replace(/<[^>]+>/g, '') : ''
          })
          .catch(() => '')
      await expect.poll(plain, { timeout: 40_000, intervals: [1_000] }).toContain(SENTENCE)
    } finally {
      await launched.app.close().catch(() => {})
    }

    const relaunched = await launchShell({
      packaged: { executable: EXE, homeDir: freshHome() },
      onboardingSeen: true,
      openFile: doc,
      settings: { starPrompt: { resolved: true } },
      videoDir: 'linux-smoke',
    })
    try {
      const docs = await waitForPageWithUrl(relaunched.app, '://docs/')
      await expect(docs.locator('.doc-page')).toContainText(SENTENCE, { timeout: 40_000 })
    } finally {
      await relaunched.app.close().catch(() => {})
    }
  })

  test('B8 a PDF page rasterizes to a non-blank canvas', async () => {
    const { PDFDocument, StandardFonts, rgb } = await import('pdf-lib')
    const dir = freshHome()
    const pdfPath = join(dir, 'sample.pdf')
    const pdfDoc = await PDFDocument.create()
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica)
    const page = pdfDoc.addPage([420, 300])
    page.drawRectangle({ x: 60, y: 70, width: 220, height: 130, color: rgb(0.8, 0.15, 0.2) })
    page.drawText('GenOffice Linux smoke', { x: 60, y: 230, size: 18, font })
    await writeFile(pdfPath, await pdfDoc.save())

    const launched = await launchShell({
      packaged: { executable: EXE, homeDir: dir },
      onboardingSeen: true,
      openFile: pdfPath,
      settings: { starPrompt: { resolved: true } },
      videoDir: 'linux-smoke',
    })
    try {
      const pdf = await waitForPageWithUrl(launched.app, '://pdf/')
      // pdf.js renders into a canvas per visible page; wasm/pdfjs payloads ship
      // under resources/wasm, so a blank canvas here means a missing payload.
      await expect.poll(() => pdf.locator('.pdf-page canvas').count(), { timeout: 40_000 }).toBeGreaterThan(0)
      const ink = await pdf.evaluate(() => {
        const canvas = document.querySelector('.pdf-page canvas') as HTMLCanvasElement
        const ctx = canvas.getContext('2d', { willReadFrequently: true })
        if (!ctx) return -1
        const { data } = ctx.getImageData(0, 0, canvas.width, canvas.height)
        let nonWhite = 0
        for (let i = 0; i < data.length; i += 4 * 89) {
          if (data[i + 3] > 0 && (data[i] < 235 || data[i + 1] < 235 || data[i + 2] < 235)) nonWhite++
        }
        return nonWhite
      })
      console.log('B8 ink samples:', ink)
      expect(ink).toBeGreaterThan(5)
      await pdf.screenshot({ path: screenshotPath('linux-B8-pdf') })
    } finally {
      await launched.app.close().catch(() => {})
    }
  })

  test('B9 Markdown and Slides open', async () => {
    // Two packaged cold boots in one test — the 90s config default is for one.
    test.setTimeout(240_000)
    const dir = freshHome()
    const md = join(dir, 'note.md')
    await writeFile(md, '# Linux smoke\n\n- [ ] check the packaged build\n')
    const deck = join(dir, 'deck.pptx')
    await writeFile(deck, '')

    const mdRun = await launchShell({
      packaged: { executable: EXE, homeDir: dir },
      onboardingSeen: true,
      openFile: md,
      settings: { starPrompt: { resolved: true } },
      videoDir: 'linux-smoke',
    })
    try {
      const page = await waitForPageWithUrl(mdRun.app, '://markdown/')
      await expect(page.locator('.doc-editor h1')).toHaveText('Linux smoke', { timeout: 40_000 })
      await page.screenshot({ path: screenshotPath('linux-B9-markdown') })
    } finally {
      await mdRun.app.close().catch(() => {})
    }

    const slideRun = await launchShell({
      packaged: { executable: EXE, homeDir: freshHome() },
      onboardingSeen: true,
      openFile: deck,
      settings: { starPrompt: { resolved: true } },
      videoDir: 'linux-smoke',
    })
    try {
      const page = await waitForPageWithUrl(slideRun.app, '://slides/')
      // The deck is drawn on a canvas (Konva), so "a stage canvas is visible"
      // is the equivalent of "the editor rendered".
      await expect(page.locator('.stage-wrap canvas').first()).toBeVisible({ timeout: 40_000 })
      await page.screenshot({ path: screenshotPath('linux-B9-slides') })
    } finally {
      await slideRun.app.close().catch(() => {})
    }
  })

  test('B10 File > Export CSV writes the grid', async () => {
    const dir = freshHome()
    const workbook = join(dir, 'csv-source.xlsx')
    const target = join(dir, 'exported.csv')
    await writeFile(workbook, await buildGridWorkbook(SMOKE_ROWS))

    const launched = await launchShell({
      packaged: { executable: EXE, homeDir: dir },
      onboardingSeen: true,
      openFile: workbook,
      settings: { starPrompt: { resolved: true } },
      videoDir: 'linux-smoke',
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
      await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
        timeout: 40_000,
      })
      await launched.app.evaluate(({ dialog }, csvPath) => {
        dialog.showSaveDialog = (async () => ({ canceled: false, filePath: csvPath })) as never
        dialog.showMessageBox = (async () => ({ response: 1, checkboxChecked: false })) as never
      }, target)
      // Re-send until it lands: the export refuses politely until preload finishes.
      await expect(
        async () => {
          await launched.app.evaluate(({ webContents }) => {
            webContents
              .getAllWebContents()
              .find((w) => w.getURL().includes('://sheets/'))
              ?.send('menu:action', 'export-csv')
          })
          expect(existsSync(target)).toBe(true)
        },
      ).toPass({ timeout: 30_000, intervals: [2_000] })
      const bytes = await readFile(target)
      expect([bytes[0], bytes[1], bytes[2]]).toEqual([0xef, 0xbb, 0xbf]) // UTF-8 BOM
      const text = bytes.subarray(3).toString('utf8')
      expect(text).toContain('项目,金额,数量,备注')
      expect(text).toContain('外委单乙,4500,12,待付款')
    } finally {
      await launched.app.close().catch(() => {})
    }
  })

  test('B11 Settings renders the AI pane without clipped rows', async () => {
    const launched = await launchShell({
      packaged: { executable: EXE, homeDir: freshHome() },
      onboardingSeen: true,
      settings: { starPrompt: { resolved: true } },
      videoDir: 'linux-smoke',
    })
    try {
      const shell = launched.page
      await shell.getByRole('button', { name: 'Settings', exact: true }).click()
      await shell.locator('.set-nav-item').filter({ hasText: 'AI Model' }).click()
      await expect(shell.locator('.set-pane-ai-model')).toBeVisible({ timeout: 20_000 })
      // Font substitution differs per platform, so the packaged Linux build is
      // exactly where a fixed-width label would start overflowing (the same
      // measurement settings-custom-endpoint-layout.spec.ts makes).
      const overflow = await shell
        .locator('.set-pane-ai-model')
        .evaluate((pane) => {
          const elements = [pane, ...pane.querySelectorAll('.set-field, .set-field-text')]
          return elements
            .filter((el) => el.scrollWidth > el.clientWidth + 1)
            .map((el) => ({ cls: el.className, w: el.clientWidth, sw: el.scrollWidth }))
        })
      expect(overflow).toEqual([])
      await shell.screenshot({ path: screenshotPath('linux-B11-settings-ai') })
    } finally {
      await launched.app.close().catch(() => {})
    }
  })
})
