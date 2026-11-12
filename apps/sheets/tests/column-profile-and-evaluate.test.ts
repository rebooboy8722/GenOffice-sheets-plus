import { describe, expect, it } from 'vitest'

import { parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'

import {
  columnProfileLine,
  executeWorkbookTool,
  textLooksNumericOrDate,
  type SheetsSkillDeps,
} from '../src/renderer/ai/tools'
import { createRangeAggregator, formatRangeAggregate } from '../src/renderer/ai/aggregate'

/**
 * P1/P2 teeth, deterministic (no model): the read path must NAME the
 * text-stored-as-number trap instead of letting the model discover it through
 * a probe formula that returns 0, and evaluate_formula must answer custom
 * computation without writing a cell.
 *
 * The motivating measurement: a real export kept its money
 * and date columns entirely as text strings; the model read them, saw
 * "15000.00", wrote SUM/COUNT, watched them return 0, and burned 5-7 turns
 * discovering what these lines state outright.
 */

describe('textLooksNumericOrDate', () => {
  it.each(['15000.00', '1,234,567', '-3.5', '12%', '2024-10-31', '2024-10-31 13:58', '2026/1/5'])(
    'flags numeric/date-shaped text: %s',
    (s) => expect(textLooksNumericOrDate(s)).toBe(true),
  )
  it.each(['north', '工程外委', '', 'A-100', '24312-0745号'])('passes plain text through: %s', (s) =>
    expect(textLooksNumericOrDate(s)).toBe(false),
  )
})

describe('columnProfileLine (read_range)', () => {
  it('names the trap column and prints the SUMPRODUCT hint', () => {
    // N2:N4 hold three numeric strings; F2:F4 hold category text
    const cells = {
      F2: { value: '劳务外委' },
      F3: { value: '工程外委' },
      F4: { value: '劳务外委' },
      N2: { value: '15000.00' },
      N3: { value: '2,000.00' },
      N4: { value: null },
    }
    const line = columnProfileLine(cells, parseRange('F2:N4'))
    expect(line).toContain('F 3/3 non-empty text')
    expect(line).toContain('N 2/3 non-empty TEXT-numeric')
    expect(line).toContain('plain SUM/COUNT/AVERAGE return 0')
    expect(line).toContain('SUMPRODUCT(VALUE(...))')
  })

  it('flags the trap even when the read includes the header row (the discovery read)', () => {
    // The first read of a sheet is typically A1:…, header included. The first
    // draft required numericText === textCells, which a header text cell
    // always defeats — the warning never fired exactly on the read that matters.
    const line = columnProfileLine(
      { N1: { value: '预估费' }, N2: { value: '15000.00' } },
      parseRange('N1:N2'),
    )
    expect(line).toContain('N 2/2 non-empty TEXT-numeric(1/2 numeric-shaped)')
    expect(line).toContain('plain SUM/COUNT/AVERAGE return 0')
  })

  it('adds no warning when everything is genuinely numeric', () => {
    const line = columnProfileLine({ B2: { value: 3 }, B3: { value: 4 } }, parseRange('B2:B3'))
    expect(line).toContain('B 2/2 non-empty numeric')
    expect(line).not.toContain('NOTE')
  })
})

describe('formatRangeAggregate (aggregate_range)', () => {
  it('states all-text when there are zero numeric cells', () => {
    const agg = createRangeAggregator()
    for (const v of ['15000.00', '2000.00', '30000.00']) agg.add(v)
    agg.addEmpty(1)
    const out = formatRangeAggregate('N2:N1258', agg.finish(5), 5)
    expect(out).toContain('stored as: TEXT')
    expect(out).toContain('numeric/date-shaped but stored as text')
    expect(out).toContain('SUMPRODUCT(VALUE(range))')
  })

  it('reports the mix when both kinds are present', () => {
    const agg = createRangeAggregator()
    agg.add(15000)
    agg.add('2000.00')
    const out = formatRangeAggregate('M2:M3', agg.finish(5), 5)
    expect(out).toContain('mixed — 1 numeric, 1 text')
  })
})

describe('evaluate_formula tool', () => {
  const deps = (_formula: string) =>
    ({
      getActiveSheetInfo: () => ({ sheetId: 'sh1', name: 'S', rows: 100, columns: 20, sheets: [] }),
      evaluateFormula: () => Promise.resolve(119151200),
    }) as unknown as SheetsSkillDeps

  it('returns the computed value and says no cell was written', async () => {
    const out = await executeWorkbookTool(
      { id: 'c1', name: 'evaluate_formula', input: { formula: '=SUMPRODUCT(VALUE(N2:N1258))' } },
      deps('=SUMPRODUCT(VALUE(N2:N1258))'),
    )
    expect(out.output).toContain('=> 119151200')
    expect(out.output).toContain('No cell was written')
    expect(out.mutated).toBe(false)
  })

  it('rejects a formula not starting with =', async () => {
    const out = await executeWorkbookTool(
      { id: 'c1', name: 'evaluate_formula', input: { formula: 'SUM(A1)' } },
      deps(''),
    )
    expect(out.output).toContain('starting with')
  })

  it('surfaces an evaluation error instead of an empty success', async () => {
    const out = await executeWorkbookTool(
      { id: 'c1', name: 'evaluate_formula', input: { formula: '=BOGUS(1)' } },
      {
        getActiveSheetInfo: () => ({ sheetId: 'sh1', name: 'S', rows: 100, columns: 20, sheets: [] }),
        evaluateFormula: () => Promise.resolve({ error: 'no value' }),
      } as unknown as SheetsSkillDeps,
    )
    expect(out.output).toContain('no value')
  })
})
