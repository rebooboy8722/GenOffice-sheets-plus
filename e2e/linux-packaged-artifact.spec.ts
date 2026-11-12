import { expect, test } from '@playwright/test'
import { createHash } from 'node:crypto'
import { execFileSync } from 'node:child_process'
import {
  existsSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

/**
 * Linux packaging verification — no Electron, no display, seconds to run.
 *
 * Why this exists: `electron-builder` **exits 0 when an extraResources source
 * is missing**. The build that produced these artifacts logged
 * `file source doesn't exist from=.../win-ocr.exe` and still reported success,
 * and its `beforePack` asserts the sidecar's presence only for win and mac.
 * So "the build passed" is not evidence that Sheets can open a workbook:
 * a missing `resources/native/xlsx-sidecar` ships green. This spec is the
 * Linux-side assertion the packaging step does not make.
 *
 * Build first (needs the 210 proxy for electron-builder's tool downloads):
 *   HTTPS_PROXY=http://127.0.0.1:7893 npm run build:all
 *   cd apps/shell && HTTPS_PROXY=... npm run dist:linux -- AppImage deb \
 *     -c.extraMetadata.version=0.12.29-fork-r3linux
 *
 * Then:  GENOFFICE_LINUX_SMOKE=1 DISPLAY= npx playwright test \
 *          --config e2e/playwright.config.ts e2e/linux-packaged-artifact.spec.ts
 *
 * These artifacts are **self-test only**: built on Ubuntu 24.04. The measured
 * glibc floor is asserted below rather than assumed (A6) — the sidecar needs
 * 2.34, which the product's documented floor (glibc >= 2.34 / Ubuntu 22.04)
 * satisfies, so the version stamp `-fork-r3linux` is what marks a local build.
 */

const ENABLED = process.platform === 'linux' && process.env.GENOFFICE_LINUX_SMOKE === '1'
const SHELL_DIR = resolve(__dirname, '../apps/shell')
const RELEASE = process.env.GENOFFICE_LINUX_DIST ?? join(SHELL_DIR, 'release')
const UNPACKED = join(RELEASE, 'linux-unpacked')
const VERSION = process.env.GENOFFICE_LINUX_VERSION ?? '0.12.29-fork-r3linux'
const SIDECAR_SOURCE = join(SHELL_DIR, '../sheets/native/xlsx-engine/target/release/xlsx-sidecar')
const APPS = ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'html'] as const

const debPath = existsSync(RELEASE)
  ? readdirSync(RELEASE).find((f) => f.endsWith('.deb'))
  : undefined
const appImagePath = existsSync(RELEASE)
  ? readdirSync(RELEASE).find((f) => f.endsWith('.AppImage'))
  : undefined
const DEB = debPath ? join(RELEASE, debPath) : ''
const APPIMAGE = appImagePath ? join(RELEASE, appImagePath) : ''
/** extracted once (no FUSE needed: --appimage-extract); reused by A5 */
const SQUASH = join(RELEASE, 'squashfs-root')

const md5 = (path: string): string =>
  createHash('md5').update(readFileSync(path)).digest('hex')

function requireArtifact(path: string, label: string): void {
  if (!existsSync(path)) {
    throw new Error(
      `missing ${label} at ${path} — build it: npm run build:all && ` +
        `(cd apps/shell && HTTPS_PROXY=http://127.0.0.1:7893 npm run dist:linux -- AppImage deb)`,
    )
  }
}

test.describe('linux packaged artifact', () => {
  test.skip(!ENABLED, 'set GENOFFICE_LINUX_SMOKE=1 (linux only)')

  test('A1 the packaged app carries the build stamp, not 0.12.0', () => {
    requireArtifact(join(UNPACKED, 'resources/app.asar'), 'linux-unpacked app.asar')
    const { extractFile } = require('@electron/asar') as {
      extractFile: (asar: string, filename: string) => Buffer
    }
    const pkg = JSON.parse(extractFile(join(UNPACKED, 'resources/app.asar'), 'package.json').toString('utf8'))
    // 0.12.0 is what package.json says — seeing it here would mean the version
    // override silently dropped, and the run would be indistinguishable from an
    // unpackaged one (which is exactly what B1's isPackaged check guards).
    expect(pkg.version).toBe(VERSION)
    expect(pkg.name).toBe('@genoffice/shell')
    expect(pkg.main).toBe('out/main/index.js')
  })

  test('A2 the xlsx sidecar is present, executable and byte-identical in all three payloads', () => {
    // The single most important assertion here: electron-builder exits 0 when
    // an extraResources source is missing, and beforePack only guards win/mac.
    // Without this, a Sheets app that cannot open a workbook builds green.
    const expected = md5(SIDECAR_SOURCE)

    const inUnpacked = join(UNPACKED, 'resources/native/xlsx-sidecar')
    requireArtifact(inUnpacked, 'sidecar in linux-unpacked')
    const head = readFileSync(inUnpacked).subarray(0, 4)
    expect([...head]).toEqual([0x7f, 0x45, 0x4c, 0x46]) // ELF magic
    expect(statSync(inUnpacked).mode & 0o111).toBeTruthy() // executable bit survived
    expect(md5(inUnpacked)).toBe(expected)

    if (DEB) {
      const listed = execFileSync('dpkg-deb', ['-c', DEB], { encoding: 'utf8', maxBuffer: 64 << 20 })
      const line = listed.split('\n').find((l) => l.includes('resources/native/xlsx-sidecar'))
      expect(line, `sidecar missing from ${debPath}`).toBeTruthy()
      // dpkg records the mode; the exec bit is what makes it runnable after install
      expect(line!.trim().startsWith('-rwx')).toBeTruthy()
      expect(line).toContain('11289328') // same size as the cargo output
    }

    if (APPIMAGE && existsSync(SQUASH)) {
      const inAppImage = join(SQUASH, 'resources/native/xlsx-sidecar')
      requireArtifact(inAppImage, 'sidecar in extracted AppImage')
      expect(md5(inAppImage)).toBe(expected)
    }
  })

  test('A3 every app module tree and the shared resources shipped', () => {
    const resources = join(UNPACKED, 'resources')
    for (const app of APPS) {
      // extraResources copies apps/<app>/out -> resources/modules/<app>; a
      // missing one means that editor tab would open a blank webContents.
      const renderer = join(resources, 'modules', app, 'renderer')
      expect(existsSync(join(renderer, 'index.html')), `modules/${app}/renderer/index.html`).toBeTruthy()
    }
    // The fork's own P0b/P1/P2 work lives in the sheets renderer bundle; pin
    // that the packaged copy is the one just built, not a stale tree.
    const source = join(SHELL_DIR, '../sheets/out/renderer/assets')
    const sourceAsset = readdirSync(source).find((f) => /^index-.*\.js$/.test(f))!
    const packagedAsset = readdirSync(join(UNPACKED, 'resources/modules/sheets/renderer/assets')).find(
      (f) => /^index-.*\.js$/.test(f),
    )!
    expect(packagedAsset).toBe(sourceAsset)
    expect(md5(join(UNPACKED, 'resources/modules/sheets/renderer/assets', packagedAsset))).toBe(
      md5(join(source, sourceAsset)),
    )

    expect(existsSync(join(resources, 'cli')), 'bundled genoffice CLI')
    expect(existsSync(join(resources, 'wasm')), 'wasm engines (pdf)')
    expect(existsSync(join(resources, 'app.asar')), 'shell main/preload')
  })

  test('A4 deb metadata is installable lineage, not a throwaway package', () => {
    test.skip(!DEB, 'no .deb in release/')
    const control = execFileSync('dpkg-deb', ['-I', DEB], { encoding: 'utf8' })
    expect(control).toContain('Package: genoffice')
    // dpkg normalizes `-` to `~` in versions; assert the stamped identity, not the spelling
    expect(control).toMatch(/Version: 0\.12\.29[~\-]fork[~\-]r3linux/)
    expect(control).toContain('Architecture: amd64')
    expect(control).toContain('libgtk-3-0')

    const listed = execFileSync('dpkg-deb', ['-c', DEB], { encoding: 'utf8', maxBuffer: 64 << 20 })
    expect(listed).toContain('/opt/GenOffice/genoffice')
    expect(listed).toContain('./usr/share/applications/genoffice.desktop')
    expect(listed).toContain('./usr/share/mime/packages/genoffice.xml')
    // the icon SET (multiple sizes), not a lone 1024px png — see the comment
    // in electron-builder.cjs about genspark-ai/genoffice#90
    const iconSizes = new Set(
      listed
        .split('\n')
        .map((l) => l.match(/hicolor\/(\d+)x\d+\/apps\/genoffice\.png/)?.[1])
        .filter((x): x is string => !!x),
    )
    expect(iconSizes.size).toBeGreaterThanOrEqual(3)
  })

  test('A5 AppImage payload matches linux-unpacked', () => {
    test.skip(!APPIMAGE, 'no .AppImage in release/')
    test.skip(!existsSync(SQUASH), 'run: ./GenOffice-*.AppImage --appimage-extract')
    const tree = (root: string): string[] => {
      const out: string[] = []
      const walk = (dir: string, prefix: string): void => {
        for (const entry of readdirSync(dir).sort()) {
          const full = join(dir, entry)
          const rel = join(prefix, entry)
          if (statSync(full).isDirectory()) walk(full, rel)
          else out.push(rel)
        }
      }
      walk(root, '')
      return out
    }
    const unpacked = tree(join(UNPACKED, 'resources'))
    const appImage = tree(join(SQUASH, 'resources'))
    // Same resource set ⇒ the AppImage is not silently dropping a payload the
    // unpacked dir has (which is the one the smoke specs drive). One known,
    // intended exception: the AppArmor profile is written by electron-builder's
    // Fpm target only (deb/rpm), and its own text scopes
    // `/opt/GenOffice/genoffice` — a path an AppImage never installs to.
    // Asserting the exception by name, so a *new* difference still fails.
    const onlyInUnpacked = unpacked.filter((p) => !appImage.includes(p))
    const onlyInAppImage = appImage.filter((p) => !unpacked.includes(p))
    expect(onlyInAppImage).toEqual([])
    expect(onlyInUnpacked).toEqual(['apparmor-profile'])
    expect(appImage.length).toBeGreaterThan(100)
  })

  test('A6 record the glibc floor (evidence for "do not ship this build")', () => {
    const floor = (binary: string): string => {
      const out = execFileSync('objdump', ['-T', binary], { encoding: 'utf8', maxBuffer: 64 << 20 })
      const versions = [...out.matchAll(/GLIBC_(\d+)\.(\d+)/g)].map((m) => Number(m[1]) * 1000 + Number(m[2]))
      const top = Math.max(...versions)
      return `${Math.floor(top / 1000)}.${top % 1000}`
    }
    const record = {
      builtOn: execFileSync('uname', ['-r'], { encoding: 'utf8' }).trim(),
      note: 'self-test artifact; the floor below is what a target machine needs',
      sidecar: floor(join(UNPACKED, 'resources/native/xlsx-sidecar')),
      mainBinary: floor(join(UNPACKED, 'genoffice')),
    }
    writeFileSync(join(UNPACKED, '..', 'linux-glibc-floor.json'), JSON.stringify(record, null, 2))
    // The product documents glibc >= 2.34 (Ubuntu 22.04). Asserting the measured
    // floor rather than assuming it: a build host cannot raise the floor above
    // what the toolchain links, and 2.34 here proves this build is NOT the
    // "24.04-only" artifact one would fear from the build host's glibc 2.39.
    expect(Number(record.sidecar.split('.')[1])).toBeLessThanOrEqual(34)
  })

  test('A7 desktop entries launch the real binary and match WM_CLASS', () => {
    // The .desktop file is generated by the Fpm target, so it exists inside the
    // deb (and the AppImage's own copy) but not in linux-unpacked. Pull just
    // that one member out of the deb payload rather than extracting 600 MB.
    test.skip(!DEB, 'no .deb in release/')
    const work = mkdtempSync(join(tmpdir(), 'genoffice-deb-'))
    try {
      execFileSync('ar', ['x', DEB], { cwd: work })
      const payload = readdirSync(work).find((f) => f.startsWith('data.tar.'))!
      const desktop = execFileSync('tar', ['-xOf', join(work, payload), './usr/share/applications/genoffice.desktop'], {
        encoding: 'utf8',
      })
      // executableName is pinned to `genoffice` in electron-builder.cjs so the
      // window's WM_CLASS and the launcher entry agree (X11 compares
      // case-sensitively, so a mismatch leaves an unlinked taskbar window).
      expect(desktop).toMatch(/^Exec=\/opt\/GenOffice\/genoffice %U$/m)
      expect(desktop).toMatch(/^StartupWMClass=genoffice$/m)
      expect(desktop).toMatch(/^Name=GenOffice$/m)
      expect(desktop).toMatch(/^Icon=genoffice$/m)
    } finally {
      rmSync(work, { recursive: true, force: true })
    }

    if (existsSync(SQUASH)) {
      const appImageDesktop = readFileSync(join(SQUASH, 'genoffice.desktop'), 'utf8')
      // AppImage launches through AppRun and electron-builder appends
      // --no-sandbox there; WM_CLASS must still agree.
      expect(appImageDesktop).toMatch(/^Exec=AppRun/m)
      expect(appImageDesktop).toMatch(/^StartupWMClass=genoffice$/m)
    }
  })
})
