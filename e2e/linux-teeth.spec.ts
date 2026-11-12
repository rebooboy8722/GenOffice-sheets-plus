import { expect, test, type Page } from '@playwright/test'
import { execFileSync } from 'node:child_process'
import { existsSync, mkdirSync, mkdtempSync, renameSync, rmSync } from 'node:fs'
import { mkdir, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { buildGridWorkbook, launchShell, screenshotPath, waitForPageWithUrl } from './helpers'

/**
 * Teeth for the Linux packaged suite: break the artifact on purpose and prove
 * the real specs go **red with a specific, attributable failure** — not "it
 * didn't crash". A green smoke suite is worth nothing if it would also pass
 * against a broken build; these tests are what say it would not.
 *
 * Every break is made in a `cp -al` hardlink copy of `linux-unpacked`: same
 * inodes, ~zero disk, and the original tree is physically untouched (renaming a
 * copy entry cannot change the original's inode). The copy lives under
 * `release/teeth/` and is rebuilt per run.
 *
 *   T-1 dead sidecar   → linux-packaged-smoke B5's assertions are load-bearing
 *   T-2 dead endpoint  → linux-packaged-ai C3's grid check can actually go red
 *   T-3 missing module → the artifact tree check (A3) and the editor both fail
 *
 *   GENOFFICE_LINUX_SMOKE=1 DISPLAY=:77 npx playwright test \
 *     --config e2e/playwright.config.ts e2e/linux-teeth.spec.ts
 */

const ENABLED = process.platform === 'linux' && process.env.GENOFFICE_LINUX_SMOKE === '1'
const SHELL_DIR = resolve(__dirname, '../apps/shell')
const RELEASE = process.env.GENOFFICE_LINUX_DIST ?? join(SHELL_DIR, 'release')
const UNPACKED = join(RELEASE, 'linux-unpacked')
const TEETH_ROOT = join(RELEASE, 'teeth')

const SMOKE_ROWS: Array<Array<string | number>> = [
  ['项目', '金额', '数量', '备注'],
  ['外委单甲', 1200, 3, '已结算'],
  ['外委单乙', 4500, 12, '待付款'],
  ['外委单丙', 880, 7, '部分结算'],
]

/**
 * Fresh hardlink copy of linux-unpacked with `mutate` applied; returns the
 * copied dir (which is what GENOFFICE_LINUX_DIST-style `EXE` paths point into).
 */
function makeBrokenCopy(name: string, mutate: (unpacked: string) => void): string {
  if (!existsSync(UNPACKED)) throw new Error(`no linux-unpacked at ${UNPACKED}`)
  const dir = join(TEETH_ROOT, name, 'linux-unpacked')
  rmSync(join(TEETH_ROOT, name), { recursive: true, force: true })
  // mkdirSync — the async `mkdir` here returned a promise nobody awaited and
  // cp then failed with ENOENT on a parent that didn't exist yet.
  mkdirSync(join(TEETH_ROOT, name), { recursive: true })
  // -a preserve attrs, -l hardlink instead of copy: ~600 MB of files share
  // inodes with the original, and mutating a *directory entry* (rename) in the
  // copy cannot touch the original file.
  execFileSync('cp', ['-al', UNPACKED, dir])
  mutate(dir)
  return dir
}

async function readCell(sheets: Page, row: number, col: number): Promise<unknown> {
  return sheets
    .evaluate(
      ([r, c]) => {
        const debug = (
          window as unknown as { __genofficeDebug?: { univerAPI: { getActiveWorkbook(): any } } }
        ).__genofficeDebug
        return debug?.univerAPI?.getActiveWorkbook()?.getActiveSheet()?.getRange(r, c, 1, 1).getValue() ?? null
      },
      [row, col],
    )
    .catch(() => null)
}

test.describe('linux packaged teeth: the suite must go red on a broken build', () => {
  test.skip(!ENABLED, 'set GENOFFICE_LINUX_SMOKE=1 (linux only)')

  test.afterAll(() => {
    rmSync(TEETH_ROOT, { recursive: true, force: true })
  })

  test('T-1 a renamed sidecar breaks the grid with an attributable error (B5 is load-bearing)', async () => {
    test.setTimeout(180_000)
    const broken = makeBrokenCopy('deadsoc', (dir) => {
      renameSync(join(dir, 'resources/native/xlsx-sidecar'), join(dir, 'resources/native/xlsx-sidecar.disabled'))
    })
    const homeDir = mkdtempSync(join(tmpdir(), 'genoffice-teeth-home-'))
    const workbook = join(homeDir, 'book.xlsx')
    await writeFile(workbook, await buildGridWorkbook(SMOKE_ROWS))

    const launched = await launchShell({
      packaged: { executable: join(broken, 'genoffice'), homeDir },
      onboardingSeen: true,
      openFile: workbook,
      settings: { starPrompt: { resolved: true } },
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
      videoDir: 'linux-teeth',
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
      // The *specific* form, observed by hand: the renderer shows the IPC
      // rejection naming the missing binary. (The main process logs
      // "XLSX sidecar failed to start"; what the user sees is this.)
      await expect
        .poll(() => sheets.evaluate(() => document.body.innerText).catch(() => ''), {
          timeout: 40_000,
          intervals: [1_000],
        })
        .toMatch(/xlsx-sidecar[\s\S]*ENOENT|ENOENT[\s\S]*xlsx-sidecar/)
      await sheets.screenshot({ path: screenshotPath('linux-T1-deadsoc') })
      // …and the data really never arrives: B5's exact assertion (A1 == 项目)
      // fails here, which is what makes B5's green mean something.
      expect(await readCell(sheets, 0, 0)).not.toBe('项目')
    } finally {
      await launched.app.close().catch(() => {})
    }
  })

  test("T-2 a dead endpoint writes nothing (C3's grid check can go red)", async () => {
    test.setTimeout(180_000)
    const homeDir = mkdtempSync(join(tmpdir(), 'genoffice-teeth-home-'))
    const workbook = join(homeDir, 'book.xlsx')
    await writeFile(workbook, await buildGridWorkbook(SMOKE_ROWS))
    const profileDir = join(homeDir, '.config', 'GenOffice')
    await mkdir(profileDir, { recursive: true })
    const settings = defaultAiSettings()
    settings.provider = 'custom'
    // Nothing listens on port 1: every request is refused. If C3's 16-cell check
    // were vacuous (old data faking green), "nothing happens" would be
    // indistinguishable from "the model wrote it" — this pins the difference.
    settings.providers.custom = { apiKey: 'dead', model: 'dead-model', baseUrl: 'http://127.0.0.1:1/v1' }
    await writeFile(join(profileDir, 'ai-settings.json'), JSON.stringify(resolveCustomEndpoints(settings)))

    const launched = await launchShell({
      packaged: { executable: join(UNPACKED, 'genoffice'), homeDir },
      onboardingSeen: true,
      openFile: workbook,
      settings: { starPrompt: { resolved: true } },
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
      videoDir: 'linux-teeth',
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
      await expect
        .poll(() => readCell(sheets, 0, 0), { timeout: 40_000, intervals: [500] })
        .toBe('项目')
      const composer = sheets.getByPlaceholder(/Ask AI to analyze/)
      await composer.click()
      await composer.fill('Write hello into B7:E10 as a 4x4 block of text. Change nothing else.')
      await sheets.getByRole('button', { name: 'Send', exact: true }).click()
      // The message must visibly reach the panel — otherwise "the grid did not
      // change" would also pass when the message never sent at all.
      await expect(sheets.getByText(/Write hello into B7/).first()).toBeVisible({ timeout: 20_000 })
      // …and with a refused endpoint, 20 s later the target block is still empty.
      await sheets.waitForTimeout(20_000)
      for (const [r, c] of [
        [6, 1],
        [6, 4],
        [9, 1],
        [9, 4],
      ]) {
        expect(await readCell(sheets, r, c), `cell ${r},${c} must stay empty`).toBeFalsy()
      }
      await sheets.screenshot({ path: screenshotPath('linux-T2-deadendpoint') })
    } finally {
      await launched.app.close().catch(() => {})
    }
  })

  test('T-3 a missing sheets module tree fails the artifact check and the editor', async () => {
    test.setTimeout(120_000)
    const broken = makeBrokenCopy('nomod', (dir) => {
      renameSync(join(dir, 'resources/modules/sheets'), join(dir, 'resources/modules/sheets.disabled'))
    })
    // The A3-style static check must catch it…
    expect(existsSync(join(broken, 'resources/modules/sheets/renderer/index.html'))).toBe(false)
    // …and so must the app: the editor webContents has nothing to load.
    const homeDir = mkdtempSync(join(tmpdir(), 'genoffice-teeth-home-'))
    const launched = await launchShell({
      packaged: { executable: join(broken, 'genoffice'), homeDir },
      onboardingSeen: true,
      settings: { starPrompt: { resolved: true } },
      videoDir: 'linux-teeth',
    })
    try {
      const shell = launched.page
      await shell.getByRole('button', { name: /AI Sheets/ }).click()
      // Two acceptable red forms: the editor view is never created (its
      // renderer file is gone), or it opens with nothing to show. Both mean the
      // smoke suite's "Sheet1 + cell content" signal can never appear — which
      // is exactly what A3 and B5 assert against.
      let sheets: Page | undefined
      try {
        sheets = await waitForPageWithUrl(launched.app, '://sheets/', 15_000)
      } catch {
        sheets = undefined
      }
      if (!sheets) {
        await shell.screenshot({ path: screenshotPath('linux-T3-nomodule') })
        return // no editor at all: the teeth are proven
      }
      await expect
        .poll(() => sheets!.evaluate(() => document.body.innerText).catch(() => ''), {
          timeout: 10_000,
          intervals: [1_000],
        })
        .not.toMatch(/Sheet1/)
      await shell.screenshot({ path: screenshotPath('linux-T3-nomodule') })
    } finally {
      await launched.app.close().catch(() => {})
    }
  })
})
