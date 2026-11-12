/// Streaming range statistics for the aggregate_range AI tool: the answer to
/// "how many distinct suppliers" on an 88k-row sheet must come from batched
/// reads, never from read_range loops (token cost) or COUNTIF-style array
/// formulas (quadratic main-thread evaluation — the app-freeze incident).

import type { CellScalar } from '@genoffice/xlsx-gateway/domain/workbook.types'

import { textLooksNumericOrDate } from './tools'

/// Distinct tracking stops (and the result says so) past this many unique
/// values; counts and numeric stats stay exact.
const MAX_DISTINCT_TRACKED = 200_000

export interface RangeAggregate {
  readonly cells: number
  readonly nonEmpty: number
  /** null when distinct tracking overflowed MAX_DISTINCT_TRACKED */
  readonly distinct: number | null
  readonly numericCount: number
  readonly sum: number
  readonly min: number | null
  readonly max: number | null
  readonly average: number | null
  /** most frequent values, descending; empty when tracking overflowed */
  readonly topValues: readonly { value: string; count: number }[]
  /** how many non-empty values are text cells; of those, how many parse as number/date */
  readonly textCount: number
  readonly numericTextCount: number
}

export interface RangeAggregator {
  add(value: CellScalar): void
  addRepeated(value: CellScalar, count: number): void
  addEmpty(count: number): void
  finish(topValueCount: number): RangeAggregate
}

export function createRangeAggregator(): RangeAggregator {
  let cells = 0
  let nonEmpty = 0
  let numericCount = 0
  let textCount = 0
  let numericTextCount = 0
  let sum = 0
  let min: number | null = null
  let max: number | null = null
  let overflowed = false
  const counts = new Map<string, number>()

  const addRepeated = (value: CellScalar, count: number): void => {
    if (count <= 0) return
    cells += count
    if (value === null || value === '') return
    nonEmpty += count
    if (typeof value === 'number') {
      numericCount += count
      sum += value * count
      min = min === null ? value : Math.min(min, value)
      max = max === null ? value : Math.max(max, value)
    } else if (typeof value === 'string') {
      // Independent of the distinct table (which is cleared on overflow):
      // "every value here is TEXT" must survive overflow, because the model's
      // next move — SUM or SUMPRODUCT(VALUE(...)) — hinges on exactly that.
      textCount += count
      if (textLooksNumericOrDate(value)) numericTextCount += count
    }
    if (overflowed) return
    const key = `${typeof value}:${String(value)}`
    const existing = counts.get(key)
    if (existing !== undefined) {
      counts.set(key, existing + count)
    } else if (counts.size >= MAX_DISTINCT_TRACKED) {
      overflowed = true
      counts.clear()
    } else {
      counts.set(key, count)
    }
  }

  return {
    add(value: CellScalar): void {
      addRepeated(value, 1)
    },
    addRepeated,
    addEmpty(count: number): void {
      if (count <= 0) return
      cells += count
    },
    finish(topValueCount: number): RangeAggregate {
      const topValues = overflowed
        ? []
        : [...counts.entries()]
            .sort((left, right) => right[1] - left[1])
            .slice(0, Math.max(0, topValueCount))
            .map(([key, count]) => ({ value: key.slice(key.indexOf(':') + 1), count }))
      return {
        cells,
        nonEmpty,
        distinct: overflowed ? null : counts.size,
        numericCount,
        sum,
        min,
        max,
        average: numericCount > 0 ? sum / numericCount : null,
        topValues,
        textCount,
        numericTextCount,
      }
    },
  }
}

const formatNumber = (value: number): string =>
  Number.isInteger(value)
    ? value.toLocaleString('en-US')
    : value.toLocaleString('en-US', { maximumFractionDigits: 4 })

/** Tool-output rendering (English, like all other tool outputs). */
export function formatRangeAggregate(
  rangeLabel: string,
  aggregate: RangeAggregate,
  topValueCount: number,
): string {
  const lines = [
    `Statistics for ${rangeLabel}:`,
    `- cells: ${formatNumber(aggregate.cells)}, non-empty: ${formatNumber(aggregate.nonEmpty)}`,
    aggregate.distinct === null
      ? `- distinct values: more than ${formatNumber(MAX_DISTINCT_TRACKED)} (tracking stopped)`
      : `- distinct values: ${formatNumber(aggregate.distinct)}`,
  ]
  if (aggregate.numericCount > 0) {
    lines.push(
      `- numeric cells: ${formatNumber(aggregate.numericCount)}, sum: ${formatNumber(aggregate.sum)}, ` +
        `average: ${formatNumber(aggregate.average ?? 0)}, min: ${formatNumber(aggregate.min ?? 0)}, max: ${formatNumber(aggregate.max ?? 0)}`,
    )
  }
  // The line the numeric line above can never print: a column whose values are
  // ALL text looks like "no numeric stats" and the model reaches for SUM —
  // which returns 0 on text. Saying it here costs one line and saves the
  // probe-and-discover loop that measured 5-7 turns on a real export.
  if (aggregate.textCount > 0 && aggregate.textCount === aggregate.nonEmpty) {
    lines.push(
      aggregate.numericTextCount === aggregate.textCount
        ? '- stored as: TEXT — every value is numeric/date-shaped but stored as text; plain SUM/COUNT/AVERAGE return 0, use SUMPRODUCT(VALUE(range)) / COUNTA / text criteria'
        : aggregate.numericTextCount > 0
          ? // Header rows land in aggregates too, so "every value numeric-shaped"
            // is too strict: say how many are, the warning is the same.
            `- stored as: TEXT — ${formatNumber(aggregate.numericTextCount)} of ${formatNumber(aggregate.textCount)} values are numeric/date-shaped text; plain SUM/COUNT/AVERAGE return 0 on those, use SUMPRODUCT(VALUE(range)) / COUNTA / text criteria`
          : '- stored as: TEXT (no numeric cells in this range)',
    )
  } else if (aggregate.textCount > 0 && aggregate.numericCount > 0) {
    lines.push(
      `- stored as: mixed — ${formatNumber(aggregate.numericCount)} numeric, ${formatNumber(aggregate.textCount)} text`,
    )
  }
  const shown = aggregate.topValues.slice(0, Math.max(0, topValueCount))
  if (shown.length > 0) {
    lines.push('- top values by frequency:')
    for (const { value, count } of shown) {
      lines.push(
        `  ${value.length > 80 ? `${value.slice(0, 80)}…` : value}: ${formatNumber(count)}`,
      )
    }
  }
  return lines.join('\n')
}
