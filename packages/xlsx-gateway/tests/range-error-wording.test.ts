import { describe, expect, it } from 'vitest'
import { z } from 'zod'

import { describeOperationErrors, workbookOperationSchema } from '../src/domain/workbook-dsl'

/**
 * What a rejected batch costs depends on what the rejection *says*. The caller
 * here is usually a model, and `range: Invalid input` — zod's default for a
 * failed regex — told it nothing, so a request phrased the way people phrase
 * them ("format column A") came back as `A:A`, got refused, and the model spent
 * a full round trip guessing the accepted shape. Measured against a local model
 * in e2e/sheets-ai-turn-count.spec.ts: one wasted turn per run, in both arms.
 */

/** Goes through the same call the app makes (propose_operations in
 * apps/sheets/src/renderer/ai/tools.ts), because the model reads this string,
 * not the raw issue list. */
function rejection(ops: unknown[]): string {
  const parsed = z.array(workbookOperationSchema).safeParse(ops)
  return parsed.success ? '' : describeOperationErrors(ops, parsed.error)
}

const format = (range: string): unknown => ({
  op: 'format_range',
  sheetId: 'sheet-1',
  range,
  format: { numberFormat: '0%' },
})

describe('range rejection wording', () => {
  it('a cell span is accepted', () => {
    expect(rejection([format('A1:A200')])).toBe('')
  })

  it('a whole-column span is refused by name, with the shape to use instead', () => {
    const text = rejection([format('A:A')])
    expect(text).toContain('whole columns or rows')
    expect(text).toContain('A1:A200')
    expect(text).not.toContain('Invalid input')
  })

  it('same for a whole-row span and a bare column letter', () => {
    expect(rejection([format('2:5')])).toContain('whole columns or rows')
    expect(rejection([format('A')])).toContain('whole columns or rows')
  })

  it('any other malformed range is told the accepted shape, not the axis hint', () => {
    const text = rejection([format('a1:b2')])
    expect(text).toContain('uppercase')
    expect(text).not.toContain('whole columns or rows')
  })

  it('the grid-limit message still wins for a span past the edge', () => {
    expect(rejection([format('A1:XFE5')])).toMatch(/grid/i)
  })

  it('a cross-sheet source ("Sheet!A1:F2343") is told to use sourceSheetId (R3 F14)', () => {
    // R2 T24: the model wrote the sheet-name prefix into source and the old
    // message only repeated the accepted shape — no pointer to sourceSheetId.
    const text = rejection([
      { op: 'copy_range', sheetId: 'sheet-2', source: '绝热工程量!A1:F2343', target: 'A1' },
    ])
    expect(text).toContain('sheet-name prefix')
    expect(text).toContain('sourceSheetId')
    expect(text).not.toContain('Invalid input')
  })

  it('the corrected form — sourceSheetId + a plain range — validates (R3 F14)', () => {
    expect(
      rejection([
        { op: 'copy_range', sheetId: 'sheet-2', sourceSheetId: 'sheet-1', source: 'A1:F2343', target: 'A1' },
      ]),
    ).toBe('')
  })

  it('a missing simple-typed field is named as missing', () => {
    const text = rejection([{ op: 'set_cell', sheetId: 'sheet-1', value: 1 }])
    expect(text).toContain('missing address (expected string)')
  })

  it('a missing union-typed field is named as missing, not "Invalid input"', () => {
    // set_cell's value is a union (string | number | boolean | null); zod v4
    // reports a missing union field as invalid_union, which used to fall into
    // the generic "value: Invalid input" branch — no hint that the field is
    // simply absent.
    const text = rejection([{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1' }])
    expect(text).toContain('missing value')
    expect(text).not.toContain('Invalid input')
  })
})
