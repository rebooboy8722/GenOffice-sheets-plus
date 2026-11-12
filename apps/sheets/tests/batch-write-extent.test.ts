import { describe, expect, it } from 'vitest'

import { workbookOperationSchema } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

import { batchWriteExtent } from '../src/renderer/write-extent'

/**
 * Teeth for the grid-growth-on-write hook.
 *
 * The hook lives in plan-operations and only fires when batchWriteExtent
 * reports a write target beyond the sheet's current grid, so every rule here
 * is load-bearing:
 *  - a write target past the edge MUST be reported (or the "out of bounds"
 *    rejection storm comes back);
 *  - a READ source must NOT be (or a bad copy source is silently papered over
 *    with a bigger grid instead of surfacing);
 *  - a structural op must NOT be (insert anchors beyond the grid are the
 *    model guessing at structure, and the error message is the correction).
 * The fixtures go through workbookOperationSchema so a field renamed upstream
 * fails here as a parse error instead of silently testing nothing.
 */

const op = (raw: Record<string, unknown>) =>
  workbookOperationSchema.parse(raw) as ReturnType<typeof workbookOperationSchema.parse>

describe('batchWriteExtent', () => {
  it('reports a single-cell write past the last column (the exact storm case)', () => {
    const got = batchWriteExtent([op({ op: 'set_formula', sheetId: 's1', address: 'AJ1', formula: '=1' })])
    expect(got.get('s1')).toEqual({ rows: 1, cols: 36 }) // AJ = column 36
  })

  it('derives set_range extent from start + values grid when range is absent', () => {
    const got = batchWriteExtent([
      op({ op: 'set_range', sheetId: 's1', start: 'AJ2', values: [[1], [2], [3], [4]] }),
    ])
    expect(got.get('s1')).toEqual({ rows: 5, cols: 36 })
  })

  it('prefers the explicit range on set_range', () => {
    const got = batchWriteExtent([op({ op: 'set_range', sheetId: 's1', range: 'AK2:AL5', values: [[1, 2]] })])
    expect(got.get('s1')).toEqual({ rows: 5, cols: 38 })
  })

  it('counts a copy/move TARGET but never its read source', () => {
    const got = batchWriteExtent([
      op({ op: 'copy_range', sheetId: 's1', source: 'A1:F100', target: 'AJ1' }),
      op({ op: 'move_range', sheetId: 's1', source: 'A2:V14', target: 'AK200' }),
    ])
    // sources (max col F/V, row 100) contribute nothing: only the targets do
    expect(got.get('s1')).toEqual({ rows: 200, cols: 37 })
  })

  it('ignores structural ops entirely — their anchor error is the correction', () => {
    const got = batchWriteExtent([
      op({ op: 'insert_cols', sheetId: 's1', column: 'AZ', count: 1 }),
      op({ op: 'insert_rows', sheetId: 's1', row: 9999, count: 2 }),
      op({ op: 'delete_cols', sheetId: 's1', column: 'A', count: 1 }),
    ])
    expect(got.size).toBe(0)
  })

  it('groups by sheetId and keeps the max per sheet', () => {
    const got = batchWriteExtent([
      op({ op: 'set_cell', sheetId: 's1', address: 'C4', value: 1 }),
      op({ op: 'set_cell', sheetId: 's1', address: 'B9', value: 2 }),
      op({ op: 'set_cell', sheetId: 's2', address: 'AA2', value: 3 }),
    ])
    expect(got.get('s1')).toEqual({ rows: 9, cols: 3 })
    expect(got.get('s2')).toEqual({ rows: 2, cols: 27 })
  })

  it('counts format_range and clear_range as writes, unknown/other ops as nothing', () => {
    const got = batchWriteExtent([
      op({ op: 'format_range', sheetId: 's1', range: 'AJ1:AJ4', format: { bold: true } }),
      op({ op: 'clear_range', sheetId: 's1', range: 'AK2' }),
      op({ op: 'sort_range', sheetId: 's1', range: 'AZ1:AZ50', byColumn: 'A', order: 'asc' }),
    ])
    // sort_range is a read+shuffle of existing data — not a write target
    expect(got.get('s1')).toEqual({ rows: 4, cols: 37 })
  })
})
