import { fsyncSync, mkdtempSync, openSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { rename } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { writeXlsxAtomically } from '../src/gateway/xlsx-gateway'
import { promoteFileAtomically } from '../src/gateway/xlsx-package-io'

vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return {
    ...actual,
    openSync: vi.fn(actual.openSync),
    fsyncSync: vi.fn(actual.fsyncSync),
  }
})

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return { ...actual, rename: vi.fn(actual.rename) }
})

let dir = ''

afterEach(() => {
  if (dir) rmSync(dir, { recursive: true, force: true })
  dir = ''
  vi.clearAllMocks()
})

/**
 * Power-loss durability of the workbook save itself.
 *
 * Sheets' .xlsx saves do **not** go through electron-utils' `atomicWriteFile` —
 * they go through the two promoters below, each with its own temp + rename. Both
 * flushed the temp file and neither flushed the **directory**, so after a crash
 * the rename could still be missing while the save had already reported success:
 * the workbook reads back with its old contents and nothing tells the user.
 * (`packages/electron-utils` got this in c8a7558a; these are the paths the sheet
 * bytes actually travel.)
 */
describe('directory flush after the rename', () => {
  const skipOnWindows = process.platform === 'win32' ? it.skip : it

  skipOnWindows('promoteFileAtomically fsyncs the directory the workbook lives in', async () => {
    dir = mkdtempSync(join(tmpdir(), 'gw-promote-'))
    const temp = join(dir, '.staged.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    writeFileSync(temp, 'workbook bytes')
    const opened: string[] = []
    const actual = await vi.importActual<typeof import('node:fs')>('node:fs')
    vi.mocked(openSync).mockImplementation((...args) => {
      opened.push(String(args[0]))
      return actual.openSync(...(args as Parameters<typeof actual.openSync>))
    })
    await promoteFileAtomically(temp, target)
    // The temp's bytes first, then the directory entry that names the workbook.
    expect(opened).toEqual([temp, dir])
    expect(vi.mocked(fsyncSync)).toHaveBeenCalledTimes(2)
    expect(readFileSync(target, 'utf8')).toBe('workbook bytes')
  })

  skipOnWindows('writeXlsxAtomically flushes the directory too', async () => {
    dir = mkdtempSync(join(tmpdir(), 'gw-write-'))
    const target = join(dir, 'book.xlsx')
    const opened: string[] = []
    const actual = await vi.importActual<typeof import('node:fs')>('node:fs')
    vi.mocked(openSync).mockImplementation((...args) => {
      opened.push(String(args[0]))
      return actual.openSync(...(args as Parameters<typeof actual.openSync>))
    })
    await writeXlsxAtomically(target, Buffer.from('bytes'))
    expect(opened[0]).not.toBe(dir) // the temp, whose name is generated inside
    expect(opened.at(-1)).toBe(dir)
    expect(dirname(target)).toBe(dir)
  })

  it('still publishes the workbook when the directory refuses to sync', async () => {
    // The rename has already landed; failing the save over a flush this
    // filesystem cannot perform would throw away the user's edit outright.
    dir = mkdtempSync(join(tmpdir(), 'gw-refuse-'))
    const temp = join(dir, '.staged.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    writeFileSync(temp, 'new')
    const actual = await vi.importActual<typeof import('node:fs')>('node:fs')
    vi.mocked(openSync).mockImplementation((...args) => {
      if (String(args[0]) === dir) {
        throw Object.assign(new Error('EIO: injected'), { code: 'EIO' })
      }
      return actual.openSync(...(args as Parameters<typeof actual.openSync>))
    })
    await expect(promoteFileAtomically(temp, target)).resolves.toBeUndefined()
    expect(readFileSync(target, 'utf8')).toBe('new')
  })

  it('does not try to open a directory handle on Windows', async () => {
    // NTFS journals the rename itself and a directory handle cannot be opened,
    // so the flush is skipped — pinned, because the skip is a platform claim.
    dir = mkdtempSync(join(tmpdir(), 'gw-win-'))
    const temp = join(dir, '.staged.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    writeFileSync(temp, 'bytes')
    const original = process.platform
    Object.defineProperty(process, 'platform', { value: 'win32', configurable: true })
    try {
      await promoteFileAtomically(temp, target)
    } finally {
      Object.defineProperty(process, 'platform', { value: original, configurable: true })
    }
    const openedDirs = vi.mocked(openSync).mock.calls.filter(([p]) => String(p) === dir)
    expect(openedDirs).toHaveLength(0)
    expect(vi.mocked(rename)).toHaveBeenCalled()
  })
})
