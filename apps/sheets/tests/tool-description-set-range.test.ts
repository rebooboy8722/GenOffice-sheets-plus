import { describe, expect, it } from 'vitest'

import { WORKBOOK_TOOLS } from '../src/renderer/ai/tools'
import { workbookOperationSchema } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

/**
 * The `set_range` fields have to be readable **without loading a guide**.
 *
 * `propose_operations` inlines a few basic ops and sends everything else behind
 * `load_guide`. `set_range` was in the "see the guide" group, and a model that
 * answers without loading one guesses: a user hit
 * `unknown field(s) startCell, data — set_range takes: sheetId, start, range, values`.
 * A batch is all-or-nothing, so that guess discarded the whole batch and burned
 * a round trip. `7f5958e3` inlined the op with its real field names.
 *
 * Why this is a unit test and not only the real-endpoint spec: whether a given
 * model run happens to send `set_range` is stochastic — with the fix reverted,
 * the model still sent it correctly on one run and dodged to `set_cell` on
 * another, so the e2e stayed green either way. What is *not* stochastic is
 * whether the description states the fields the validator accepts, and that is
 * what this pins. The example is parsed by the real schema rather than compared
 * to a hardcoded field list, so renaming a field on either side turns this red
 * instead of letting the description silently drift from the validator.
 */

const description = WORKBOOK_TOOLS.find((t) => t.name === 'propose_operations')?.description ?? ''

describe('propose_operations description inlines set_range', () => {
  it('names set_range with its real fields, not behind load_guide', () => {
    // The inline section runs up to "load_guide before using them"; anything
    // after that marker is a guide the model must spend a turn loading.
    const [inlinePart, guidePart = ''] = description.split('load_guide before using them')
    expect(inlinePart).toContain('set_range')
    // set_range appears in the guide list too (harmless, it is genuinely there),
    // so the assertion that matters is that it is *also* in the inline part.
    expect(guidePart).toContain('writing(')
  })

  it('uses field names the validator actually accepts', () => {
    // Exactly the shape the description shows the model.
    const parsed = workbookOperationSchema.safeParse({
      op: 'set_range',
      sheetId: 'sheet-1',
      start: 'B2',
      values: [['x', 'y']],
    })
    expect(parsed.success).toBe(true)
  })

  it('rejects the invented field names the description warns about', () => {
    // The description tells the model "NOT startCell/data"; if the validator ever
    // starts accepting those, that warning is wrong and this test should say so.
    expect(
      workbookOperationSchema.safeParse({
        op: 'set_range',
        sheetId: 'sheet-1',
        startCell: 'B2',
        data: [['x', 'y']],
      }).success,
    ).toBe(false)
  })

  it('keeps the warning text that names the wrong guesses', () => {
    // The user-visible failure came from a model that had never seen the words
    // `start`/`values` next to `set_range`; naming the wrong guesses explicitly
    // is what makes a second occurrence unlikely.
    expect(description).toMatch(/startCell/)
    expect(description).toMatch(/values is a 2D array/)
  })
})
