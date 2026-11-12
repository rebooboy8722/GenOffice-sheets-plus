import { parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { RangeBounds } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

/**
 * The farthest row/column a batch's WRITE targets reach, per sheet.
 *
 * An imported sheet's grid is exactly its data extent (`MINIMUM_SHEET_*` in
 * univer-sync only floor tiny sheets), so "put the result in the first free
 * column to the right" addresses a cell that does not exist — and the model,
 * having no legal way to write there, was measured burning 31 "Range is out of
 * bounds" rejections and then inserting a column *inside* the user's data to
 * manufacture one. The fix is to grow the grid to the
 * write target before validation, exactly what grid-growth does when a human
 * scrolls to the edge. This module computes what "far enough" is.
 *
 * Only write targets count. Read sources (copy/move `source`, fill `source`,
 * chart `dataRange`) and structural ops stay out deliberately: a read past the
 * data is a model error that should surface, not something to paper over with
 * a bigger grid, and an insert anchor beyond the grid is the model guessing at
 * structure. Whitelist, not blacklist: an op type not listed here contributes
 * nothing, so new ops default to today's behaviour.
 */

/** 1-based row/column counts (what getMaxRows/getMaxColumns return). */
export interface SheetWriteExtent {
  rows: number
  cols: number
}

const boundsToExtent = (b: RangeBounds): SheetWriteExtent => ({
  rows: b.endRow + 1,
  cols: b.endColumn + 1,
})

/** `set_range` without `range`: the top-left `start` plus the values grid. */
const startPlusValuesExtent = (op: {
  start?: string | undefined
  values: unknown[][]
}): RangeBounds | null => {
  if (!op.start) return null // neither field — existing validation rejects it
  const anchor = parseRange(op.start)
  const rows = op.values.length
  const cols = Math.max(...op.values.map((row) => row.length))
  return {
    startRow: anchor.startRow,
    startColumn: anchor.startColumn,
    endRow: anchor.startRow + rows - 1,
    endColumn: anchor.startColumn + cols - 1,
  }
}

export function batchWriteExtent(
  operations: readonly WorkbookOperation[],
): Map<string, SheetWriteExtent> {
  const out = new Map<string, SheetWriteExtent>()
  const push = (sheetId: string, extent: SheetWriteExtent): void => {
    const prev = out.get(sheetId)
    if (!prev || extent.rows > prev.rows || extent.cols > prev.cols) {
      out.set(sheetId, {
        rows: Math.max(prev?.rows ?? 0, extent.rows),
        cols: Math.max(prev?.cols ?? 0, extent.cols),
      })
    }
  }
  for (const operation of operations) {
    let bounds: RangeBounds | null
    switch (operation.op) {
      case 'set_cell':
      case 'set_formula':
      case 'clear_cell':
        bounds = parseRange(operation.address)
        break
      case 'set_range':
        bounds = operation.range
          ? parseRange(operation.range)
          : startPlusValuesExtent(operation)
        break
      case 'clear_range':
      case 'format_range':
        bounds = parseRange(operation.range)
        break
      // Target only: the source is a read and must NOT grow the grid.
      case 'fill_range':
      case 'copy_range':
      case 'move_range':
        bounds = parseRange(operation.target)
        break
      default:
        continue // reads, structural ops, unknown ops: today's behaviour
    }
    if (bounds) push(operation.sheetId, boundsToExtent(bounds))
  }
  return out
}
