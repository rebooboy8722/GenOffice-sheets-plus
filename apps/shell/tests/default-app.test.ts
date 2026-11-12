import { describe, expect, it, vi } from 'vitest'
import {
  OFFICE_TYPES,
  createDefaultAppService,
  decodeRegOutput,
  macAppBundlePath,
  parseMacStatus,
  parseRegValue,
} from '../src/main/default-app'
import type { RunCommand } from '../src/main/default-app'

const ME = 'com.genoffice.app'
const WPS = 'com.kingsoft.wpsoffice.mac'

function macStatusJson(
  owner: (uti: string) => string | null,
  name: string | null = 'WPS Office',
): string {
  return JSON.stringify({
    me: ME,
    types: OFFICE_TYPES.map((t) => {
      const id = owner(t.uti)
      return { uti: t.uti, id, name: id === ME ? 'GenOffice' : id ? name : null }
    }),
  })
}

describe('macAppBundlePath', () => {
  it('walks up from the executable to the .app bundle', () => {
    expect(macAppBundlePath('/Applications/GenOffice.app/Contents/MacOS/GenOffice')).toBe(
      '/Applications/GenOffice.app',
    )
    expect(macAppBundlePath('/usr/local/bin/electron')).toBeNull()
  })
})

describe('parseMacStatus', () => {
  it('is default when every type resolves to our bundle id', () => {
    expect(parseMacStatus(macStatusJson(() => ME))).toEqual({
      state: 'default',
      others: [],
      manualOnly: false,
    })
  })

  it('names the other app once even when it owns several types', () => {
    expect(parseMacStatus(macStatusJson(() => WPS))).toEqual({
      state: 'other',
      others: ['WPS Office'],
      manualOnly: false,
    })
  })

  it('is "other" when a single type belongs to someone else', () => {
    const json = macStatusJson((uti) => (uti.includes('presentationml') ? WPS : ME))
    expect(parseMacStatus(json).state).toBe('other')
  })

  it('falls back to the bundle id when the owner has no display name', () => {
    expect(parseMacStatus(macStatusJson(() => 'com.example.x', null)).others).toEqual([
      'com.example.x',
    ])
  })

  it('treats a type with no handler as not yet claimed', () => {
    const json = macStatusJson((uti) => (uti.includes('macroenabled') ? null : ME))
    expect(parseMacStatus(json)).toEqual({ state: 'other', others: [], manualOnly: false })
  })

  it('is unknown when LaunchServices has no handler at all', () => {
    expect(parseMacStatus(macStatusJson(() => null)).state).toBe('unknown')
  })
})

describe('parseRegValue', () => {
  it('reads the data column of a reg query line', () => {
    const out = [
      '',
      'HKEY_CURRENT_USER\\Software\\...\\.docx\\UserChoice',
      '    ProgId    REG_SZ    Word.Document.12',
      '',
    ].join('\r\n')
    expect(parseRegValue(out)).toBe('Word.Document.12')
    expect(parseRegValue('    (Default)    REG_SZ    Word Document\r\n')).toBe('Word Document')
    expect(
      parseRegValue('ERROR: The system was unable to find the specified registry key'),
    ).toBeNull()
  })
})

describe('decodeRegOutput', () => {
  it('keeps ASCII output untouched (ProgIds are ASCII)', () => {
    const line = '    ProgId    REG_SZ    Word.Document.12\r\n'
    expect(decodeRegOutput(Buffer.from(line, 'utf8'))).toBe(line)
  })

  it('recovers GBK (Chinese system ACP) CJK names that fail strict UTF-8', () => {
    // 'WPS \u8868\u683C' in GBK: 57 50 53 20 B1 ED B8 F1 (invalid UTF-8 \u2192 fallback)
    expect(decodeRegOutput(Buffer.from('57505320B1EDB8F1', 'hex'))).toBe('WPS \u8868\u683C')
  })

  it('passes genuine UTF-8 output through the strict path', () => {
    const line = '    (Default)    REG_SZ    WPS \u8868\u683C\r\n'
    expect(decodeRegOutput(Buffer.from(line, 'utf8'))).toBe(line)
  })
})

describe('createDefaultAppService', () => {
  const base = {
    packaged: true,
    exePath: '/Applications/GenOffice.app/Contents/MacOS/GenOffice',
    openExternal: vi.fn(async () => {}),
  }

  it('is unsupported in dev builds and never runs a command', async () => {
    const run = vi.fn<RunCommand>()
    const svc = createDefaultAppService({ ...base, platform: 'darwin', packaged: false, run })
    expect((await svc.status()).state).toBe('unsupported')
    expect((await svc.set()).state).toBe('unsupported')
    expect(run).not.toHaveBeenCalled()
  })

  it('mac: claims every UTI through osascript and re-reads the status', async () => {
    let owner = WPS
    const run = vi.fn<RunCommand>(async (cmd, args) => {
      expect(cmd).toBe('osascript')
      expect(args.slice(0, 3)).toEqual(['-l', 'JavaScript', '-e'])
      expect(args[4]).toBe('/Applications/GenOffice.app')
      expect(args.slice(5)).toEqual(OFFICE_TYPES.map((t) => t.uti))
      if (args[3].includes('LSSetDefaultRoleHandlerForContentType')) {
        owner = ME
        return JSON.stringify({ me: ME, failed: [] })
      }
      return macStatusJson(() => owner)
    })
    const svc = createDefaultAppService({ ...base, platform: 'darwin', run })
    expect(await svc.status()).toMatchObject({ state: 'other', others: ['WPS Office'] })
    expect(await svc.set()).toMatchObject({ state: 'default', others: [] })
    expect(run).toHaveBeenCalledTimes(3)
  })

  it('linux: compares xdg-mime answers with our desktop id', async () => {
    const calls: string[][] = []
    const run = vi.fn<RunCommand>(async (cmd, args) => {
      calls.push([cmd, ...args])
      if (args[0] === 'query')
        return args[2].includes('spreadsheetml') ? 'wps-office-et.desktop\n' : 'genoffice.desktop\n'
      return ''
    })
    const svc = createDefaultAppService({
      ...base,
      platform: 'linux',
      run,
      readFile: (p) => {
        if (p.endsWith('/usr/share/applications/wps-office-et.desktop'))
          return '[Desktop Entry]\nName=WPS Spreadsheets\n'
        throw new Error('ENOENT')
      },
    })
    expect(await svc.status()).toEqual({
      state: 'other',
      others: ['WPS Spreadsheets'],
      manualOnly: false,
    })
    await svc.set()
    expect(calls.find((c) => c[1] === 'default')).toEqual([
      'xdg-mime',
      'default',
      'genoffice.desktop',
      ...OFFICE_TYPES.map((t) => t.mime),
    ])
  })

  it('windows: reg query drives the status, nothing but reg.exe is spawned', async () => {
    const run = vi.fn<RunCommand>()
    const runRaw = vi.fn(async (_cmd: string, args: string[]) => {
      const key = args[1]
      if (key.includes('UserChoice')) {
        if (key.includes('.pptx')) throw new Error('no user choice')
        return Buffer.from(
          key.includes('.docx')
            ? '    ProgId    REG_SZ    Word.Document.12\r\n'
            : '    ProgId    REG_SZ    Excel Workbook\r\n',
          'utf8',
        )
      }
      if (key === 'HKCR\\.pptx')
        return Buffer.from('    (Default)    REG_SZ    PowerPoint Presentation\r\n', 'utf8')
      if (key === 'HKCR\\Word.Document.12')
        return Buffer.from('    (Default)    REG_SZ    Microsoft Word Document\r\n', 'utf8')
      throw new Error('missing')
    })
    const openExternal = vi.fn(async () => {})
    const svc = createDefaultAppService({ ...base, platform: 'win32', run, runRaw, openExternal })
    expect(await svc.status()).toEqual({
      state: 'other',
      others: ['Microsoft Word Document'],
      manualOnly: true,
    })
    expect(run).not.toHaveBeenCalled() // the powershell probe is gone entirely
    await svc.set()
    expect(openExternal).toHaveBeenCalledWith('ms-settings:defaultapps')
  })

  it('windows: GBK display names from reg.exe survive the code-page round trip', async () => {
    const run = vi.fn<RunCommand>()
    // reg.exe on a Chinese (GBK 936) system emits CJK display names as GBK
    // bytes (hex literals — TextEncoder cannot produce GBK): WPS 文字 / 表格 / 演示
    const GBK: Record<string, string> = {
      'WPS 文字': '57505320CEC4D7D6',
      'WPS 表格': '57505320B1EDB8F1',
      'WPS 演示': '57505320D1DDCABE',
    }
    const runRaw = vi.fn(async (_cmd: string, args: string[]) => {
      const key = args[1]
      if (key.includes('UserChoice')) throw new Error('no user choice')
      const cls: Record<string, string> = {
        'HKCR\\.docx': 'WPS.Doc',
        'HKCR\\.xlsx': 'WPS.Xls',
        'HKCR\\.xlsm': 'WPS.Xls',
        'HKCR\\.xls': 'WPS.Xls',
        'HKCR\\.pptx': 'WPS.Ppt',
      }
      if (key in cls) return Buffer.from(`    (Default)    REG_SZ    ${cls[key]}\r\n`, 'utf8')
      const names: Record<string, string> = {
        'HKCR\\WPS.Doc': 'WPS 文字',
        'HKCR\\WPS.Xls': 'WPS 表格',
        'HKCR\\WPS.Ppt': 'WPS 演示',
      }
      if (key in names) {
        // prefix + GBK name + CRLF, byte-concatenated (the name is NOT UTF-8)
        return Buffer.concat([
          Buffer.from('    (Default)    REG_SZ    ', 'utf8'),
          Buffer.from(GBK[names[key]], 'hex'),
          Buffer.from('\r\n', 'utf8'),
        ])
      }
      throw new Error('missing')
    })
    const svc = createDefaultAppService({ ...base, platform: 'win32', run, runRaw })
    expect(await svc.status()).toEqual({
      state: 'other',
      others: ['WPS 文字', 'WPS 表格', 'WPS 演示'],
      manualOnly: true,
    })
    expect(run).not.toHaveBeenCalled()
  })

  it('reports unknown instead of throwing when the probe fails', async () => {
    const run = vi.fn<RunCommand>(async () => {
      throw new Error('osascript missing')
    })
    const svc = createDefaultAppService({ ...base, platform: 'darwin', run })
    expect((await svc.status()).state).toBe('unknown')
  })
})
