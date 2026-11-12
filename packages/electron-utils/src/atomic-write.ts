import { randomBytes } from 'node:crypto'
import {
  closeSync,
  fsyncSync,
  mkdirSync,
  openSync,
  renameSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs'
import { copyFile, open, rename, unlink, writeFile } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'

/** Transient Windows codes: antivirus/indexer briefly locks the rename target. */
const RETRYABLE_RENAME_CODES = new Set(['EPERM', 'EACCES', 'EBUSY'])
const RENAME_RETRY_DELAYS_MS = [50, 100, 200, 400]

/**
 * Codes a refused fsync may surface with inside cloud-sync folders, under AV
 * locks or on filesystems without fsync. The bytes are already written, so a
 * refused flush only weakens power-loss durability and must not fail the save.
 */
const TOLERATED_SYNC_CODES = new Set(['EPERM', 'EACCES', 'EBUSY', 'EINVAL', 'ENOSYS'])

const errorCode = (error: unknown) => (error as NodeJS.ErrnoException).code ?? ''
const isRetryableRename = (error: unknown) => RETRYABLE_RENAME_CODES.has(errorCode(error))
const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))
const sleepSync = (ms: number) => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms)

const tempPathBeside = (filePath: string) =>
  join(dirname(filePath), `.${basename(filePath)}.${randomBytes(6).toString('hex')}.tmp`)

async function syncFileBestEffort(path: string): Promise<void> {
  let handle
  try {
    handle = await open(path, 'r+')
  } catch (error) {
    if (TOLERATED_SYNC_CODES.has(errorCode(error))) return
    throw error
  }
  try {
    await handle.sync()
  } catch (error) {
    if (!TOLERATED_SYNC_CODES.has(errorCode(error))) throw error
  } finally {
    await handle.close()
  }
}

/**
 * Flushes the directory so a completed rename survives power loss.
 *
 * `rename` is atomic but only *in the filesystem's terms*: after a crash the
 * directory entry can still point at the old inode while the temp file's data
 * blocks are on disk, which reads back as "I saved it and the file is the old
 * one" — the save reported success, so nothing tells the user otherwise.
 *
 * Skipped on Windows, where NTFS journals the rename itself and opening a
 * directory handle is not supported. Every failure is tolerated: on filesystems
 * without fsync, in cloud-sync folders, or under AV locks the rename has already
 * happened, and refusing the save then would lose the user's work outright to
 * protect against a risk we cannot remove.
 */
async function syncDirBestEffort(dirPath: string): Promise<void> {
  if (process.platform === 'win32') return
  let handle
  try {
    handle = await open(dirPath, 'r')
  } catch {
    return
  }
  try {
    await handle.sync()
  } catch {
    /* see above: the rename already landed */
  } finally {
    await handle.close().catch(() => {})
  }
}

function syncDirBestEffortSync(dirPath: string): void {
  if (process.platform === 'win32') return
  let descriptor: number
  try {
    descriptor = openSync(dirPath, 'r')
  } catch {
    return
  }
  try {
    fsyncSync(descriptor)
  } catch {
    /* see syncDirBestEffort */
  } finally {
    try {
      closeSync(descriptor)
    } catch {
      /* already closed */
    }
  }
}

function syncFileBestEffortSync(path: string): void {
  let descriptor: number
  try {
    descriptor = openSync(path, 'r+')
  } catch (error) {
    if (TOLERATED_SYNC_CODES.has(errorCode(error))) return
    throw error
  }
  try {
    fsyncSync(descriptor)
  } catch (error) {
    if (!TOLERATED_SYNC_CODES.has(errorCode(error))) throw error
  } finally {
    closeSync(descriptor)
  }
}

/**
 * Same-dir temp file, best-effort fsync, then rename. Rename-over-existing fails
 * transiently on Windows (Defender/indexer locks: retry) and permanently on some
 * network mounts (gvfs SMB refuses with EEXIST, genoffice#1877); either way the
 * save falls back to an in-place write rather than failing.
 */
export async function atomicWriteFile(filePath: string, data: Buffer): Promise<void> {
  const tmp = tempPathBeside(filePath)
  try {
    await writeFile(tmp, data)
    await syncFileBestEffort(tmp)
  } catch (error) {
    await unlink(tmp).catch(() => {})
    throw error
  }
  try {
    await renameWithRetry(tmp, filePath)
    await syncDirBestEffort(dirname(filePath))
    return
  } catch {
    // fall through to the in-place write
  }
  // Keep the completed temp until the fallback lands: if that write fails
  // or the process dies, the new bytes still exist somewhere on disk.
  await writeFile(filePath, data)
  // The fallback overwrote the file's contents in place, so the file — not a
  // directory entry — is what needs flushing. A refused flush (the tolerated
  // codes) passes silently; a real I/O error throws, deliberately: the old
  // contents are already gone, so claiming success would be the worse failure.
  // The temp then stays behind on purpose — it is the only surviving copy of the
  // new bytes, which is what the rule above this write is for.
  await syncFileBestEffort(filePath)
  await unlink(tmp).catch(() => {})
}

async function renameWithRetry(tmp: string, filePath: string): Promise<void> {
  for (let attempt = 0; ; attempt += 1) {
    try {
      await rename(tmp, filePath)
      return
    } catch (error) {
      if (!isRetryableRename(error) || attempt >= RENAME_RETRY_DELAYS_MS.length) throw error
      await sleep(RENAME_RETRY_DELAYS_MS[attempt] ?? 0)
    }
  }
}

/** copyFile stays in the kernel, so a multi-hundred-MB file is never read into RAM */
export async function atomicCopyFile(source: string, filePath: string): Promise<void> {
  const tmp = tempPathBeside(filePath)
  try {
    await copyFile(source, tmp)
    await syncFileBestEffort(tmp)
  } catch (error) {
    await unlink(tmp).catch(() => {})
    throw error
  }
  try {
    await renameWithRetry(tmp, filePath)
    await syncDirBestEffort(dirname(filePath))
    return
  } catch {
    // fall through to the in-place write
  }
  // Keep the completed temp until the fallback lands, as atomicWriteFile does.
  await copyFile(source, filePath)
  await syncFileBestEffort(filePath)
  await unlink(tmp).catch(() => {})
}

/**
 * Atomic replacement for small userData JSON state files (recents, stars, AI
 * settings). Same retry + in-place fallback as atomicWriteFile; the retry
 * sleeps block the caller, which is acceptable for a few KB under a transient
 * Windows lock.
 */
export function writeJsonAtomic(filePath: string, value: unknown): void {
  mkdirSync(dirname(filePath), { recursive: true })
  const json = JSON.stringify(value, null, 2)
  const tmp = tempPathBeside(filePath)
  const removeTemp = () => {
    try {
      unlinkSync(tmp)
    } catch {
      /* temp never created */
    }
  }
  try {
    writeFileSync(tmp, json, { encoding: 'utf8', flag: 'wx' })
    syncFileBestEffortSync(tmp)
    for (let attempt = 0; ; attempt += 1) {
      try {
        renameSync(tmp, filePath)
        syncDirBestEffortSync(dirname(filePath))
        return
      } catch (error) {
        if (!isRetryableRename(error) || attempt >= RENAME_RETRY_DELAYS_MS.length) throw error
        sleepSync(RENAME_RETRY_DELAYS_MS[attempt] ?? 0)
      }
    }
  } catch (error) {
    if (isRetryableRename(error)) {
      writeFileSync(filePath, json, 'utf8')
      // Same as atomicWriteFile's fallback: this path replaced the file's
      // contents rather than a directory entry, so the file needs the flush. A
      // real I/O error throws — the previous JSON is already gone — and the
      // temp below stays behind as the only surviving copy of the new bytes.
      syncFileBestEffortSync(filePath)
      removeTemp()
      return
    }
    removeTemp()
    throw error
  }
}
