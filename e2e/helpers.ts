/**
 * Shared launcher for Electron E2E tests.
 *
 * Each test boots the built shell (`apps/shell/out`) against a scratch
 * userData dir (via GENOFFICE_USER_DATA) so runs never touch real settings
 * and never collide with a running install's single-instance lock.
 * Build first: `npm run build:all`.
 */
import { _electron as electron, type ElectronApplication, type Page } from '@playwright/test'
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { createRequire } from 'node:module'

export const SHELL_DIR = resolve(__dirname, '../apps/shell')
export const ARTIFACTS_DIR = resolve(__dirname, 'artifacts')

const SHELL_MAIN = join(SHELL_DIR, 'out/main/index.js')

interface LaunchOptions {
  /** reuse a previous scratch dir to simulate a second launch */
  userDataDir?: string
  /**
   * Launch a **packaged** build (`release/linux-unpacked/genoffice`) instead of
   * `apps/shell/out` under the repo's Electron.
   *
   * `homeDir` is a scratch directory used as `HOME`/`XDG_*` for the run. It has
   * to be HOME rather than `GENOFFICE_USER_DATA` because every consumer of that
   * variable gates it on `!app.isPackaged` (`apps/shell/src/main/index.ts:428`,
   * `sheets-main.ts:5000`, `slides-main.ts:5185`) — a packaged app ignores it,
   * and the point of these tests is the packaged code path. Electron derives
   * userData and the single-instance lock from HOME, so redirecting HOME
   * isolates settings *and* the lock with no main-process change; the resulting
   * profile is `$HOME/.config/GenOffice`, which is what the product itself
   * resolves (`packages/cli/src/gui.ts`).
   *
   * `userDataDir` (returned and pre-seeded) is then that profile directory, so
   * the existing `app-settings.json` seeding works unchanged.
   */
  packaged?: { executable: string; homeDir: string }
  /** UI language override (GENOFFICE_LANG); defaults to English for stable assertions */
  lang?: string
  /** pre-seed app-settings.json with onboardingSeen=true to start at the home screen */
  onboardingSeen?: boolean
  /** extra app-settings.json keys (e.g. defaultSaveDir) written before launch */
  settings?: Record<string, unknown>
  /** subdir of e2e/artifacts to store this launch's video in */
  videoDir: string
  /** absolute document path passed as argv, opened in an editor tab on launch */
  openFile?: string
  /** extra environment variables for the launched app */
  env?: Record<string, string>
}

export interface LaunchedApp {
  app: ElectronApplication
  page: Page
  userDataDir: string
}

export async function launchShell(options: LaunchOptions): Promise<LaunchedApp> {
  const packaged = options.packaged
  if (packaged) {
    if (!existsSync(packaged.executable)) {
      throw new Error(`Missing packaged executable at ${packaged.executable} — run \`npm run dist:linux\` first`)
    }
  } else if (!existsSync(SHELL_MAIN)) {
    throw new Error(`Missing build output at ${SHELL_MAIN} — run \`npm run build:all\` first`)
  }
  // A packaged run gets its profile from HOME, so the profile dir is derived
  // rather than scratch-named; it is still what gets seeded and returned.
  const profileDir = packaged ? join(packaged.homeDir, '.config', 'GenOffice') : undefined
  const userDataDir =
    options.userDataDir ?? profileDir ?? (await mkdtemp(join(tmpdir(), 'genoffice-e2e-')))
  if (packaged) await mkdir(packaged.homeDir, { recursive: true })
  await mkdir(userDataDir, { recursive: true })
  if (options.onboardingSeen || options.settings) {
    await writeFile(
      join(userDataDir, 'app-settings.json'),
      JSON.stringify({
        ...(options.onboardingSeen ? { onboardingSeen: true } : {}),
        ...options.settings,
      }),
    )
  }
  // A packaged binary is its own Electron and takes only document paths —
  // passing the app dir would be read as a file to open.
  const executablePath = packaged
    ? packaged.executable
    : (createRequire(join(SHELL_DIR, 'package.json'))('electron') as unknown as string)
  // ELECTRON_RUN_AS_NODE (set by VS Code/CI hosts) would boot Electron as
  // plain Node with no windows — strip it so the app always starts as an app
  const { ELECTRON_RUN_AS_NODE: _electronRunAsNode, ...hostEnv } = process.env
  // Linux CI runners restrict unprivileged user namespaces (no usable SUID
  // sandbox) and run under xvfb without GPU — without these the window opens
  // but the renderer never loads. The suite drives trusted local builds only.
  // Switches go before the app path so Chromium is guaranteed to consume them
  // and they never leak into the argv the app parses for documents to open.
  const args: string[] = []
  if (process.platform === 'linux') args.push('--no-sandbox', '--disable-gpu')
  if (!packaged) args.push(SHELL_DIR)
  if (options.openFile) args.push(options.openFile)
  // A packaged app ignores GENOFFICE_USER_DATA (gated on !app.isPackaged), so
  // isolation comes from redirecting HOME and the XDG dirs into the scratch home.
  const homeEnv: Record<string, string> = packaged
    ? {
        HOME: packaged.homeDir,
        XDG_CONFIG_HOME: join(packaged.homeDir, '.config'),
        XDG_CACHE_HOME: join(packaged.homeDir, '.cache'),
        XDG_DATA_HOME: join(packaged.homeDir, '.local', 'share'),
      }
    : { GENOFFICE_USER_DATA: userDataDir }
  const app = await electron.launch({
    executablePath,
    args,
    env: {
      ...hostEnv,
      ...homeEnv,
      GENOFFICE_NO_SPARE_VIEW: '1',
      GENOFFICE_LANG: options.lang ?? 'en',
      ...(options.env ?? {}),
      ...(process.platform === 'linux' ? { ELECTRON_DISABLE_SANDBOX: '1' } : {}),
    },
    // Playwright's Electron screencast wedges the page CDP session on Linux
    // (page.url() stays empty, no lifecycle events, evaluate hangs) — record
    // only where it works
    recordVideo:
      process.platform === 'linux'
        ? undefined
        : {
            dir: join(ARTIFACTS_DIR, options.videoDir),
            size: { width: 1280, height: 800 },
          },
  })
  const page = await app.firstWindow()
  await waitForDocumentReady(app, page)
  return { app, page, userDataDir }
}

/**
 * Playwright can attach to the Electron window mid-navigation and miss the
 * load lifecycle events entirely (Linux timing) — waitForLoadState then hangs
 * on a page that is actually loaded. Polling through evaluate uses the live
 * CDP session instead of the missed events.
 */
async function waitForDocumentReady(
  app: ElectronApplication,
  page: Page,
  timeoutMs = 30_000,
): Promise<void> {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    // the pre-navigation about:blank document also reports readyState complete
    const ready = await page
      .evaluate(() =>
        document.readyState !== 'loading' && window.location.href !== 'about:blank'
          ? document.readyState
          : null,
      )
      .catch(() => null)
    if (ready) return
    await new Promise((r) => setTimeout(r, 100))
  }
  // process list tells renderer-spawn failures apart from slow loads
  const diag = await app
    .evaluate(({ app: electronApp, BrowserWindow }) => ({
      processes: electronApp.getAppMetrics().map((m) => m.type),
      contents: BrowserWindow.getAllWindows().map((w) => w.webContents.getURL()),
    }))
    .catch((e) => String(e))
  throw new Error(`Shell window never loaded (url: ${page.url()}, diag: ${JSON.stringify(diag)})`)
}

/**
 * Close the app and return the recorded video path for the given page.
 *
 * Open editor tabs trigger a native Save/Don't Save/Cancel dialog on close,
 * which would block app.close() forever — stub the dialog to answer
 * "Don't Save" (button index 1) so shutdown stays unattended. If close still
 * hangs, force-kill the process after 20s and wait for close to finish.
 */
export async function closeAndSaveVideo(
  launched: LaunchedApp,
  name: string,
): Promise<string | undefined> {
  const video = launched.page.video()
  await launched.app
    .evaluate(({ dialog }) => {
      dialog.showMessageBox = (async () => ({
        response: 1,
        checkboxChecked: false,
      })) as typeof dialog.showMessageBox
    })
    .catch(() => {})
  const child = launched.app.process()
  const killTimer = setTimeout(() => {
    // SIGTERM can enter Electron's graceful quit path and leave it alive.
    // child.killed only means a signal was sent, not that the process exited.
    if (child.exitCode === null && child.signalCode === null) child.kill('SIGKILL')
  }, 20_000)
  try {
    await launched.app.close()
  } finally {
    clearTimeout(killTimer)
  }
  if (!video) return undefined
  const target = join(ARTIFACTS_DIR, 'videos', `${name}.webm`)
  try {
    await video.saveAs(target)
    return target
  } catch {
    return undefined
  }
}

export function screenshotPath(name: string): string {
  return join(ARTIFACTS_DIR, 'screenshots', `${name}.png`)
}

/**
 * Wait for a page whose URL contains `urlPart` (e.g. an editor WebContentsView).
 * Checks windows that already exist before listening, so it never races the
 * view being created between launch and the first waitForEvent call.
 */
export async function waitForPageWithUrl(
  app: ElectronApplication,
  urlPart: string,
  timeoutMs = 30_000,
): Promise<Page> {
  const deadline = Date.now() + timeoutMs
  for (;;) {
    for (const candidate of app.windows()) {
      if (candidate.url().includes(urlPart)) return candidate
      // page.url() stays empty when attach raced navigation; ask the document
      const href = await candidate.evaluate(() => window.location.href).catch(() => '')
      if (href.includes(urlPart)) return candidate
    }
    const remaining = deadline - Date.now()
    if (remaining <= 0) throw new Error(`No window with URL containing "${urlPart}"`)
    await app.waitForEvent('window', { timeout: Math.min(remaining, 1_000) }).catch(() => {})
  }
}

/**
 * Viewport point at the centre of a grid cell, for clicks a real user would
 * make (row/column headers, a cell to focus before a shortcut).
 *
 * Univer draws the sheet on a canvas, so there is no element to click: the
 * geometry is the only address. ~46px of row header, ~24px of column header,
 * 74×23 default cells. `column: -1` is the row header.
 */
export async function gridCellCenter(
  page: Page,
  row: number,
  column: number,
): Promise<{ x: number; y: number }> {
  const grid = await page.evaluate(() => {
    for (const canvas of document.querySelectorAll('canvas')) {
      const r = canvas.getBoundingClientRect()
      if (r.width > 500 && r.height > 300) return { x: r.x, y: r.y }
    }
    return null
  })
  if (!grid) throw new Error('worksheet canvas not found')
  return { x: grid.x + 46 + 37 + column * 74, y: grid.y + 24 + 12 + row * 23 }
}

/**
 * A one-sheet xlsx holding `rows`, written as a file rather than typed into the
 * grid — because for an imported workbook the AI's overview reports the extent
 * and contents of the **file on disk**, so unsaved grid edits are invisible to
 * the model. A test that seeds through the grid and then asks about "column A"
 * measures a model facing an empty sheet instead of the task it means to ask.
 */
export async function buildGridWorkbook(rows: Array<Array<string | number>>): Promise<Buffer> {
  const { default: JSZip } = await import('jszip')
  const columnLabel = (index: number): string => {
    let label = ''
    let remaining = index + 1
    while (remaining > 0) {
      remaining -= 1
      label = String.fromCharCode(65 + (remaining % 26)) + label
      remaining = Math.floor(remaining / 26)
    }
    return label
  }
  const sheetXml = rows
    .map((row, rowIndex) => {
      const cells = row
        .map((value, colIndex) => {
          const ref = `${columnLabel(colIndex)}${rowIndex + 1}`
          return typeof value === 'number'
            ? `<c r="${ref}"><v>${value}</v></c>`
            : `<c r="${ref}" t="inlineStr"><is><t>${String(value)}</t></is></c>`
        })
        .join('')
      return `<row r="${rowIndex + 1}">${cells}</row>`
    })
    .join('')
  const lastColumn = columnLabel(Math.max(0, ...rows.map((r) => r.length)) - 1)
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs>
</styleSheet>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:${lastColumn}${rows.length}"/>
  <sheetData>${sheetXml}</sheetData>
</worksheet>`,
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}
