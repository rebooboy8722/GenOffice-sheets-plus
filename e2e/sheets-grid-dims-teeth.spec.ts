import { expect, test } from '@playwright/test'
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, gridCellCenter, waitForPageWithUrl } from './helpers'

/**
 * Teeth for the collateral-damage check in `sheets-ai-large-real-table.spec.ts`.
 *
 * That spec compares the sheet's grid dimensions before and after an AI run and
 * fails if the grid grew, because the run it is meant to catch inserted a column,
 * wrote its working formulas into the new column, restored the displaced header
 * text, and answered one column further right — leaving every cell of A..AI
 * byte-identical. A content-only fingerprint called that "unchanged".
 *
 * An assertion nobody has seen fail is not evidence, though — that is exactly the
 * mistake this fork has already made twice with model-behaviour tests. So this
 * spec drives the structural change through the *human* affordance (ribbon →
 * Insert column), with no model in the loop, and asserts the dimension read
 * notices it. If this passes, `dimsChanged` is a real detector; if it stays green
 * while the AI spec's dims never move, the detector is blind and the AI spec's
 * green runs prove nothing.
 *
 * Needs the real export (same file as the AI spec) and the built shell.
 */

const SRC = process.env.GENOFFICE_E2E_XLSX ?? ''

interface DebugWindow {
  __genofficeDebug: {
    univerAPI: {
      getActiveWorkbook(): {
        getActiveSheet(): {
          getMaxRows(): number
          getMaxColumns(): number
        } | null
      } | null
    }
  }
}

test('the grid-dimension fingerprint notices a column inserted by hand', async () => {
  test.skip(!SRC, 'set GENOFFICE_E2E_XLSX to the real export (kept out of git)')
  test.setTimeout(300_000)

  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-dims-'))
  const workbook = join(scratch, 'book.xlsx')
  await copyFile(SRC, workbook)
  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-e2e-'))

  const launched = await launchShell({
    userDataDir,
    onboardingSeen: true,
    openFile: workbook,
    videoDir: 'bigtable',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const sheets = await waitForPageWithUrl(launched.app, '://sheets/', 120_000)
    await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 120_000,
    })
    // 'Sheet1' in the tab label is not "the grid is up": on a 3.4 MB sheet the
    // canvas appears seconds later, and the first run of this spec died in
    // gridCellCenter with `worksheet canvas not found`. Wait for the same canvas
    // the helper looks for (the criteria are the helper's own).
    await sheets.waitForFunction(
      () =>
        [...document.querySelectorAll('canvas')].some((c) => {
          const r = c.getBoundingClientRect()
          return r.width > 500 && r.height > 300
        }),
      null,
      { timeout: 120_000 },
    )

    const dims = () =>
      sheets.evaluate(() => {
        const ws = (window as unknown as DebugWindow).__genofficeDebug.univerAPI
          .getActiveWorkbook()
          ?.getActiveSheet()
        return { rows: ws?.getMaxRows() ?? -1, cols: ws?.getMaxColumns() ?? -1 }
      })

    // Polled like the AI spec does: a 3.4 MB sheet is still hydrating for a few
    // seconds after the page appears, and a read inside that window returns a
    // partial sheet (the first run of the sibling spec saw A1 empty, B1:G1 full).
    let before = await dims()
    for (let i = 0; i < 30 && !(before.rows > 1000 && before.cols > 30); i++) {
      await sheets.waitForTimeout(1_000)
      before = await dims()
    }
    expect(before, 'the export should open as a 1258 x 35 grid').toEqual({ rows: 1258, cols: 35 })

    // The human path: click column A's header, then the ribbon's Insert column.
    const header = await gridCellCenter(sheets, -1, 0)
    await sheets.mouse.click(header.x, header.y)
    await sheets.waitForTimeout(400)
    await sheets.getByRole('button', { name: 'Insert column', exact: true }).click()
    await sheets.waitForTimeout(1_500)

    const after = await dims()
    expect(after.rows, 'inserting a column must not change the row count').toBe(before.rows)
    expect(after.cols, 'the fingerprint must notice a column being inserted').toBe(before.cols + 1)
  } finally {
    await launched.app.close().catch(() => {})
  }
})
