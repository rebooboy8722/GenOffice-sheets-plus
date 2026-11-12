/**
 * Lightweight formula evaluator for the demo workbook (convo-sim F3).
 *
 * The demo has no Univer engine, so `evaluate_formula` needs a stand-in.
 * This covers the function set the model actually reaches for on a sparse
 * sheet (R1: 16× SUMPRODUCT((col="x")*col), SUMIF/COUNTIF/COUNTIFS with
 * "<>" / "" / wildcard criteria, SUM, COUNTA, bare cross-sheet refs) and
 * answers everything else with a clear "not supported in the demo" error
 * naming the fallback — never a silent wrong number.
 *
 * Supported: =SUM(ranges) =SUMIF(range,crit[,sum]) =SUMIFS(sum,r,c,...)
 *            =COUNTIF(range,crit) =COUNTIFS(r,c,...) =SUMPRODUCT(factors)
 *            =COUNTA/COUNT/AVERAGE/MAX/MIN/COUNTBLANK(ranges)
 *            =MATCH(lookup,range[,type]) =LARGE(range,k)  =Sheet!A1
 * Ranges may be cells (A1), rectangles (A1:B2) or whole columns (A:A).
 * SUMPRODUCT factors: a range, (range op literal) with op = < > <= >= <>,
 * a number literal, ISNUMBER(SEARCH("text",range)) as a contains-mask, or
 * ROW(range) as the row numbers — with an optional leading "--" (double
 * negation) and scalar broadcast, matching Excel's array semantics.
 * Non-numeric cells count as 0.
 *
 * F13 (R2): 46 of R2's 78 tool errors were evaluate_formula failures caused
 * by this stand-in being thinner than the real Univer engine the app uses —
 * the model reached for ISNUMBER(SEARCH(...)) masks, ROW sums, MAX/COUNTBLANK
 * and got parse errors the production app would never produce. The additions
 * below close the idioms R2 actually used; deeper nesting (IF/VALUE/...)
 * still errors, and the error names the working fallback instead.
 */

export type DemoCellValue = string | number | boolean | null
export type DemoCellMap = Readonly<
  Record<string, { readonly value: DemoCellValue; readonly formula?: string } | undefined>
>

export type SheetResolver = (sheetKey: string | undefined) => DemoCellMap | undefined

export type DemoFormulaResult = DemoCellValue | { error: string }

const SUPPORTED = [
  'SUM',
  'SUMIF',
  'SUMIFS',
  'COUNTIF',
  'COUNTIFS',
  'SUMPRODUCT',
  'COUNTA',
  'COUNT',
  'AVERAGE',
  'MAX',
  'MIN',
  'COUNTBLANK',
  'MATCH',
  'LARGE',
]

const CELL_RE = /^\$?([A-Z]{1,3})\$?(\d+)$/
const COL_RE = /^[A-Z]{1,3}$/

function colIndex(letters: string): number {
  let n = 0
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}

function colLabel(index: number): string {
  let n = index
  let out = ''
  while (n > 0) {
    const r = (n - 1) % 26
    out = String.fromCharCode(65 + r) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

type CellRef =
  | { sheetKey: string | undefined; kind: 'cells'; addresses: string[] }
  | { sheetKey: string | undefined; kind: 'columns'; c1: number; c2: number }

/** Parse `Sheet!A1` / `A1` / `A:B` / `A1:B2`; full rows error. */
function parseRef(raw: string): { ref: CellRef | undefined; error?: string } {
  let sheetKey: string | undefined
  let body = raw.trim().replace(/\$/g, '')
  const bang = body.lastIndexOf('!')
  if (bang >= 0) {
    const sheet = body.slice(0, bang).trim()
    body = body.slice(bang + 1).trim()
    if (!/^[A-Za-z_一-龥][\w一-龥]*$|^sheet-\d+$/.test(sheet)) {
      return { ref: undefined, error: `cannot parse sheet reference "${raw}"` }
    }
    sheetKey = sheet
  }
  const parts = body.split(':')
  if (parts.length === 1) {
    const cell = (parts[0] ?? '').match(CELL_RE)
    if (!cell) return { ref: undefined, error: `cannot parse cell reference "${raw}"` }
    return { ref: { sheetKey, kind: 'cells', addresses: [`${cell[1]}${cell[2]}`] } }
  }
  if (parts.length !== 2) return { ref: undefined, error: `cannot parse range "${raw}"` }
  const a = parts[0] ?? ''
  const b = parts[1] ?? ''
  // whole columns: A:A / A:B
  if (COL_RE.test(a) && COL_RE.test(b)) {
    return { ref: { sheetKey, kind: 'columns', c1: colIndex(a), c2: colIndex(b) } }
  }
  const cellA = a.match(CELL_RE)
  const cellB = b.match(CELL_RE)
  if (!cellA || !cellB) {
    return { ref: undefined, error: `cannot parse range "${raw}" (whole rows like 2:5 are not supported)` }
  }
  const r1 = Math.min(+cellA[2]!, +cellB[2]!)
  const r2 = Math.max(+cellA[2]!, +cellB[2]!)
  const c1 = Math.min(colIndex(cellA[1]!), colIndex(cellB[1]!))
  const c2 = Math.max(colIndex(cellA[1]!), colIndex(cellB[1]!))
  if ((r2 - r1 + 1) * (c2 - c1 + 1) > 1_000_000) {
    return { ref: undefined, error: `range ${raw} is too large for the demo evaluator` }
  }
  const addresses: string[] = []
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) addresses.push(`${colLabel(c)}${r}`)
  return { ref: { sheetKey, kind: 'cells', addresses } }
}

function unquote(crit: string): string {
  const c = crit.trim()
  if (c.length >= 2 && c.startsWith('"') && c.endsWith('"')) return c.slice(1, -1)
  return c
}

/** Excel-style criteria match: "" empty, "<>" non-empty, rel-ops, * / ? wildcards. */
export function matchCriteria(value: DemoCellValue, critRaw: string): boolean {
  const crit = unquote(critRaw)
  if (crit === '') return value === null || value === ''
  if (crit === '<>') return value !== null && value !== ''
  const rel = crit.match(/^(<=|>=|<|>)(.*)$/)
  if (rel) {
    const n = typeof value === 'number' ? value : NaN
    const m = Number(crit.slice(rel[1]!.length))
    if (Number.isNaN(n) || Number.isNaN(m)) return false
    switch (rel[1]) {
      case '<': return n < m
      case '>': return n > m
      case '<=': return n <= m
      case '>=': return n >= m
    }
  }
  if (crit.includes('*') || crit.includes('?')) {
    const escaped = crit.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const re = new RegExp(`^${escaped.replace(/\\\*/g, '.*').replace(/\\\?/g, '.')}$`, 'i')
    return re.test(String(value ?? ''))
  }
  return String(value ?? '').toLowerCase() === crit.toLowerCase()
}

function numeric(v: DemoCellValue): number {
  return typeof v === 'number' ? v : 0
}

/** Resolve a ref to live cell values (sparse map: absent = empty). */
function valuesFor(ref: CellRef, resolve: SheetResolver): DemoCellValue[] {
  const map = resolve(ref.sheetKey)
  if (!map) {
    throw new Error(
      ref.sheetKey ? `unknown sheet "${ref.sheetKey}"` : 'no active sheet available',
    )
  }
  if (ref.kind === 'cells') {
    return ref.addresses.map((a) => map[a]?.value ?? null)
  }
  // whole columns: enumerate the present cells, row-major
  const wanted = new Set<number>()
  for (let c = ref.c1; c <= ref.c2; c++) wanted.add(c)
  return Object.entries(map)
    .map(([address, cell]) => {
      if (!cell) return null
      const m = address.match(CELL_RE)
      return m ? { col: colIndex(m[1]!), row: +m[2]!, value: cell.value } : null
    })
    .filter((e): e is NonNullable<typeof e> => e !== null && wanted.has(e.col))
    .sort((a, b) => a.row - b.row || a.col - b.col)
    .map((e) => e.value)
}

/** Split on `sep` at parenthesis depth 0 (outside quotes) — so the comma in
 * SEARCH("m³",D2:D5) inside SUMPRODUCT(arg1,arg2) never splits the args. */
function splitTopLevel(input: string, sep: string): string[] {
  const out: string[] = []
  let cur = ''
  let inStr = false
  let depth = 0
  for (const ch of input) {
    if (ch === '"') inStr = !inStr
    if (!inStr) {
      if (ch === '(') depth += 1
      else if (ch === ')') depth = Math.max(0, depth - 1)
      else if (ch === sep && depth === 0) {
        out.push(cur)
        cur = ''
        continue
      }
    }
    cur += ch
  }
  out.push(cur)
  return out
}

/** Split a parenthesized `range op literal` (operators found outside quotes). */
function splitComparison(expr: string): { range: string; op: string; crit: string } | null {
  let inStr = false
  for (let i = 0; i < expr.length; i++) {
    const ch = expr[i]
    if (ch === '"') {
      inStr = !inStr
      continue
    }
    if (inStr) continue
    for (const op of ['<=', '>=', '<>', '=', '<', '>']) {
      if (expr.startsWith(op, i)) {
        return { range: expr.slice(0, i).trim(), op, crit: expr.slice(i + op.length).trim() }
      }
    }
  }
  return null
}

type Factor =
  | { kind: 'range'; ref: CellRef }
  | { kind: 'compare'; ref: CellRef; op: string; crit: string }
  | { kind: 'literal'; value: number }
  /** ISNUMBER(SEARCH("text",range)) → 1 where the cell's text contains `text`. */
  | { kind: 'contains'; ref: CellRef; text: string }
  /** ROW(range) → the row numbers of the range's cells. */
  | { kind: 'rownum'; ref: CellRef }

function parseFactor(raw: string): { factor?: Factor | undefined; error?: string | undefined } {
  let body = raw.trim()
  // Excel's double negation coerces TRUE/FALSE masks to 1/0 — the model writes it.
  if (body.startsWith('--')) body = body.slice(2).trim()
  // Function-call factors — bare calls, no outer parens (the forms R2 actually used).
  // ISNUMBER(SEARCH("m³",range)) — the text-contains mask idiom (R2 T14)
  const search = body.match(/^ISNUMBER\(\s*SEARCH\(\s*"([^"]*)"\s*,\s*(.+?)\s*\)\s*\)$/i)
  if (search) {
    const { ref, error } = parseRef(search[2]!)
    if (!ref) return { error }
    return { factor: { kind: 'contains', ref, text: search[1]! } }
  }
  // ROW(range) — row numbers as a factor (R2 T7: sum of blank-cell row numbers)
  const row = body.match(/^ROW\(\s*(.+?)\s*\)$/i)
  if (row) {
    const { ref, error } = parseRef(row[1]!)
    if (!ref) return { error }
    if (ref.kind === 'columns') {
      return { error: `ROW() over whole columns is not supported in the demo — use a row-bounded range like ROW(F2:F2313)` }
    }
    return { factor: { kind: 'rownum', ref } }
  }
  if (/^\(.*\)$/.test(body)) {
    const comp = splitComparison(body.slice(1, -1).trim())
    if (comp) {
      const { ref, error } = parseRef(comp.range)
      if (!ref) return { error }
      return { factor: { kind: 'compare', ref, op: comp.op, crit: unquote(comp.crit) } }
    }
  }
  if (/^-?\d+(\.\d+)?$/.test(body)) return { factor: { kind: 'literal', value: Number(body) } }
  const { ref, error } = parseRef(body)
  if (!ref) {
    // A function call we do not compute: say what IS supported, name the
    // fallback — R2's "cannot parse cell reference \"ISNUMBER(SEARCH(…\"" gave
    // the model nothing to act on and it burned 6 turns guessing.
    if (body.includes('(')) {
      return {
        error:
          `SUMPRODUCT factor "${body}" uses a nested function the demo evaluator does not compute ` +
          '(it handles range masks like (D2:D5="x"), ISNUMBER(SEARCH("x",D2:D5)) and ROW(F2:F5)). ' +
          'For other logic use SUMIF/COUNTIF, or extract the matching rows with copy_range\'s filterColumn + filterValues.',
      }
    }
    return { error }
  }
  return { factor: { kind: 'range', ref } }
}

function factorValues(factor: Factor, resolve: SheetResolver): number[] {
  if (factor.kind === 'literal') return [factor.value]
  if (factor.kind === 'contains') {
    const text = factor.text.toLowerCase()
    return valuesFor(factor.ref, resolve).map((v) =>
      v === null || v === '' ? 0 : String(v).toLowerCase().includes(text) ? 1 : 0,
    )
  }
  if (factor.kind === 'rownum') {
    // parseFactor rejects whole-column ROW; here the ref is cells-only
    const addresses = factor.ref.kind === 'cells' ? factor.ref.addresses : []
    return addresses.map((a) => +((a.match(CELL_RE) ?? [])[2] ?? NaN))
  }
  const values = valuesFor(factor.ref, resolve)
  if (factor.kind === 'range') return values.map(numeric)
  if (factor.op === '=' || factor.op === '<>') {
    const eq = values.map((v) => matchCriteria(v, factor.crit))
    return eq.map((hit) => (factor.op === '=' ? (hit ? 1 : 0) : hit ? 0 : 1))
  }
  // numeric rel-ops
  return values.map((v) => {
    const n = typeof v === 'number' ? v : NaN
    const m = Number(factor.crit)
    if (Number.isNaN(n) || Number.isNaN(m)) return 0
    switch (factor.op) {
      case '<': return (n < m ? 1 : 0)
      case '>': return (n > m ? 1 : 0)
      case '<=': return (n <= m ? 1 : 0)
      case '>=': return (n >= m ? 1 : 0)
      default: return 0
    }
  })
}

/** True if `chars` appears at parenthesis depth 0 outside quotes.
 * `skipLeading` exempts the first char (a unary minus on a bare negative). */
function hasTopLevelChar(expr: string, chars: string, skipLeading: boolean): boolean {
  let depth = 0
  let inStr = false
  for (let i = 0; i < expr.length; i++) {
    const ch = expr.charAt(i)
    if (ch === '"') inStr = !inStr
    if (inStr) continue
    if (ch === '(') depth += 1
    else if (ch === ')') depth = Math.max(0, depth - 1)
    else if (depth === 0 && (skipLeading ? i > 0 : true) && chars.includes(ch)) return true
  }
  return false
}

export function evaluateDemoFormula(
  formula: string,
  resolve: SheetResolver,
): DemoFormulaResult {
  let expr = formula.trim()
  if (expr.startsWith('=')) expr = expr.slice(1).trim()
  if (!expr) return { error: 'empty formula' }

  // R12 (候选③ 第四次): top-level binary operators / string concatenation
  // used to fall through into argument-level "cannot parse range/cell" errors
  // (misleading — the model reshuffled the form and burned calls: R10 T15 ×2,
  // R11 T1 &-concat ×5, R12 T3 ×2, R12 T4 LARGE-concat ×2). Name the real
  // problem (single function call) and the actionable path instead.
  if (hasTopLevelChar(expr, '&', false)) {
    return {
      error:
        `string concatenation with & ("${formula}") is not supported by the demo evaluator — ` +
        'it computes a single function call or a bare cell reference. ' +
        'Use evaluate_formula per part, or read the cells and concatenate in your answer.',
    }
  }
  if (hasTopLevelChar(expr, '+-*/', true)) {
    return {
      error:
        `binary arithmetic ("${formula}") is not supported by the demo evaluator — ` +
        'it computes a single function call or a bare cell reference. ' +
        'Call evaluate_formula separately for each part (e.g. =SUM(F2:F100) and =SUM(F101:F200)), ' +
        'or use aggregate_range.',
    }
  }

  // bare cell reference: =A1 / =汇总!C74
  if (!expr.includes('(')) {
    const { ref, error } = parseRef(expr)
    if (!ref) {
      return {
        error:
          `cannot parse "${formula}" — the demo evaluator only handles simple function calls ` +
          'over cell ranges (or a single cell like =A1)',
      }
    }
    let values: DemoCellValue[]
    try {
      values = valuesFor(ref, resolve)
    } catch (e) {
      // unknown sheet 等 resolve 失败必须落成 error 结果，不许抛穿
      return { error: e instanceof Error ? e.message : String(e) }
    }
    return values.length === 1 ? (values[0] ?? null) : { error: `expected a single cell, got "${expr}"` }
  }

  const call = expr.match(/^([A-Za-z]+)\s*\((.*)\)$/s)
  if (!call) {
    return {
      error:
        `cannot parse "${formula}" — the demo evaluator only understands simple single-function ` +
        'formulas over cell ranges (e.g. =SUM(F2:F33), =SUMIF(B2:B20000,"m³",F2:F20000), ' +
        '=SUMPRODUCT((D2:D20000="x")*F2:F20000)). For anything else use aggregate_range or read the cells.',
    }
  }
  const name = call[1]!.toUpperCase()
  if (!SUPPORTED.includes(name)) {
    return {
      error:
        `"${call[1]}" is not supported by the demo formula evaluator (supported: ${SUPPORTED.join(', ')}). ` +
        'Use aggregate_range for plain statistics, or read the cells and compute in your answer. ' +
        'To extract the ROWS that match a condition, use propose_operations\' copy_range with filterColumn + filterValues — ' +
        'do not retry the same formula with small variations.',
    }
  }
  const args = splitTopLevel(call[2] ?? '', ',').map((a) => a.trim())
  const parseArg = (i: number): { ref: CellRef | undefined; error?: string } => parseRef(args[i] ?? '')

  try {
    switch (name) {
      case 'SUM':
      case 'COUNTA':
      case 'COUNT':
      case 'AVERAGE': {
        if (args.length < 1) return { error: `${name} needs at least one range` }
        // F21: no `push(...spread)` — V8 caps variadic args at 65,535, so a
        // single range with more cells (e.g. A1:Z5000 = 130k) threw
        // "Maximum call stack size exceeded". parts.flat() is iterative.
        const parts: DemoCellValue[][] = []
        for (let i = 0; i < args.length; i++) {
          const { ref, error } = parseArg(i)
          if (!ref) return { error: error ?? `cannot parse "${args[i]}"` }
          parts.push(valuesFor(ref, resolve))
        }
        const values = parts.flat()
        if (name === 'SUM') return values.reduce<number>((acc, v) => acc + numeric(v), 0)
        if (name === 'COUNTA') return values.filter((v) => v !== null && v !== '').length
        if (name === 'COUNT') return values.filter((v) => typeof v === 'number').length
        const nums = values.filter((v): v is number => typeof v === 'number')
        if (nums.length === 0) return { error: 'AVERAGE: the ranges contain no numeric cells' }
        return nums.reduce((a, b) => a + b, 0) / nums.length
      }
      case 'SUMIF': {
        if (args.length < 2 || args.length > 3) {
          return { error: 'SUMIF takes 2 or 3 arguments: range, criteria[, sum_range]' }
        }
        const { ref, error } = parseArg(0)
        if (!ref) return { error: error ?? `cannot parse "${args[0]}"` }
        const values = valuesFor(ref, resolve)
        const crit = unquote(args[1] ?? '')
        let sum = 0
        if (args.length === 3) {
          const { ref: sRef, error: sErr } = parseArg(2)
          if (!sRef) return { error: sErr ?? `cannot parse sum range "${args[2]}"` }
          const sumValues = valuesFor(sRef, resolve)
          if (sumValues.length !== values.length) {
            return { error: 'SUMIF: criteria range and sum range have different sizes' }
          }
          for (let i = 0; i < values.length; i++) {
            if (matchCriteria(values[i] ?? null, crit)) sum += numeric(sumValues[i] ?? null)
          }
        } else {
          for (const v of values) if (matchCriteria(v, crit)) sum += numeric(v)
        }
        return sum
      }
      case 'SUMIFS': {
        if (args.length < 3 || args.length % 2 === 0) {
          return { error: 'SUMIFS takes the sum range first, then pairs of (range, criteria)' }
        }
        const { ref: sumRef, error } = parseArg(0)
        if (!sumRef) return { error: error ?? `cannot parse "${args[0]}"` }
        const sumValues = valuesFor(sumRef, resolve)
        const pairs: Array<{ values: DemoCellValue[]; crit: string }> = []
        for (let i = 1; i < args.length; i += 2) {
          const { ref, error: pErr } = parseArg(i)
          if (!ref) return { error: pErr ?? `cannot parse "${args[i]}"` }
          const values = valuesFor(ref, resolve)
          if (values.length !== sumValues.length) return { error: 'SUMIFS: all ranges must have the same size' }
          pairs.push({ values, crit: unquote(args[i + 1] ?? '') })
        }
        let sum = 0
        for (let i = 0; i < sumValues.length; i++) {
          if (pairs.every((p) => matchCriteria(p.values[i] ?? null, p.crit))) sum += numeric(sumValues[i] ?? null)
        }
        return sum
      }
      case 'COUNTIF': {
        if (args.length !== 2) return { error: 'COUNTIF takes 2 arguments: range, criteria' }
        const { ref, error } = parseArg(0)
        if (!ref) return { error: error ?? `cannot parse "${args[0]}"` }
        const crit = unquote(args[1] ?? '')
        return valuesFor(ref, resolve).filter((v) => matchCriteria(v, crit)).length
      }
      case 'COUNTIFS': {
        if (args.length < 2 || args.length % 2 !== 0) {
          return {
            error:
              `COUNTIFS takes pairs of (range, criteria) — got ${args.length} argument(s), an odd count. ` +
              'For a single condition use COUNTIF(range, criteria).',
          }
        }
        const arrays: Array<{ values: DemoCellValue[]; crit: string }> = []
        for (let i = 0; i < args.length; i += 2) {
          const { ref, error: pErr } = parseArg(i)
          if (!ref) return { error: pErr ?? `cannot parse "${args[i]}"` }
          arrays.push({ values: valuesFor(ref, resolve), crit: unquote(args[i + 1] ?? '') })
        }
        const n = arrays[0]!.values.length
        if (arrays.some((a) => a.values.length !== n)) {
          return { error: 'COUNTIFS: all ranges must have the same size' }
        }
        let count = 0
        for (let i = 0; i < n; i++) {
          if (arrays.every((a) => matchCriteria(a.values[i] ?? null, a.crit))) count++
        }
        return count
      }
      case 'MAX':
      case 'MIN': {
        if (args.length < 1) return { error: `${name} needs at least one range` }
        // F21: same fix as the SUM/COUNTA branch above — flat() instead of
        // push(...spread), which overflows past 65,535 variadic elements.
        const parts: DemoCellValue[][] = []
        for (let i = 0; i < args.length; i++) {
          const { ref, error } = parseArg(i)
          if (!ref) return { error: error ?? `cannot parse "${args[i]}"` }
          parts.push(valuesFor(ref, resolve))
        }
        const values = parts.flat()
        let best = NaN
        for (const v of values) {
          if (typeof v !== 'number') continue
          if (Number.isNaN(best) || (name === 'MAX' ? v > best : v < best)) best = v
        }
        if (Number.isNaN(best)) return { error: `${name}: the ranges contain no numeric cells` }
        return best
      }
      case 'COUNTBLANK': {
        if (args.length < 1) return { error: 'COUNTBLANK needs at least one range' }
        let blanks = 0
        for (let i = 0; i < args.length; i++) {
          const { ref, error } = parseArg(i)
          if (!ref) return { error: error ?? `cannot parse "${args[i]}"` }
          for (const v of valuesFor(ref, resolve)) if (v === null || v === '') blanks++
        }
        return blanks
      }
      case 'MATCH': {
        if (args.length < 2) return { error: 'MATCH takes (lookup_value, range[, match_type])' }
        const { ref, error } = parseArg(1)
        if (!ref) return { error: error ?? `cannot parse "${args[1]}"` }
        const lookup = args[0] ?? ''
        const values = valuesFor(ref, resolve)
        if (lookup.includes('(')) {
          // R9/R10 T12 ×10: "not found (exact match only)" read like a data
          // problem and the model retried with swapped arguments — the lookup
          // is a formula the demo cannot compute, say so and name the path.
          return {
            error:
              `MATCH: the lookup value is a formula ("${lookup}") the demo cannot compute — ` +
              'find the matching row with find_cells, or use LARGE(range,k) for the k-th largest value.',
          }
        }
        const idx = values.findIndex((v) => v !== null && matchCriteria(v, lookup))
        if (idx < 0) return { error: `MATCH: ${lookup} not found in the range (exact match only)` }
        return idx + 1
      }
      case 'LARGE': {
        if (args.length < 2) return { error: 'LARGE takes (range, k)' }
        const { ref, error } = parseArg(0)
        if (!ref) return { error: error ?? `cannot parse "${args[0]}"` }
        const k = Number(args[1])
        if (!Number.isInteger(k) || k < 1) return { error: 'LARGE: k must be a positive integer' }
        const nums = valuesFor(ref, resolve)
          .filter((v): v is number => typeof v === 'number')
          .sort((a, b) => b - a)
        if (k > nums.length) return { error: `LARGE: k=${k} exceeds the ${nums.length} numeric cells` }
        return nums[k - 1]!
      }
      case 'SUMPRODUCT': {
        if (args.length < 1) return { error: 'SUMPRODUCT needs at least one factor' }
        const argValues: number[][] = []
        for (const arg of args) {
          let values: number[] = []
          for (const part of splitTopLevel(arg, '*')) {
            const { factor, error } = parseFactor(part)
            if (!factor) return { error: error ?? `cannot parse SUMPRODUCT factor "${part}"` }
            const fv = factorValues(factor, resolve)
            if (values.length === 0) values = fv
            else if (fv.length === 1 && values.length > 1) {
              // scalar factor (a number literal) broadcasts, Excel-style
              const k = fv[0] ?? 0
              for (let i = 0; i < values.length; i++) values[i] = (values[i] ?? 0) * k
            } else if (values.length === 1 && fv.length > 1) {
              const k = values[0] ?? 0
              values = fv.map((x) => x * k)
            } else {
              if (fv.length !== values.length) return { error: 'SUMPRODUCT: factors have different sizes' }
              for (let i = 0; i < values.length; i++) values[i] = (values[i] ?? 0) * (fv[i] ?? 0)
            }
          }
          argValues.push(values)
        }
        const n = argValues[0]!.length
        if (argValues.some((v) => v.length !== n)) {
          return { error: 'SUMPRODUCT: all arguments must cover the same number of cells' }
        }
        let total = 0
        for (let i = 0; i < n; i++) {
          let prod = 1
          for (const v of argValues) prod *= v[i] ?? 0
          total += prod
        }
        return total
      }
    }
  } catch (e) {
    return { error: e instanceof Error ? e.message : String(e) }
  }
  return { error: `unhandled function ${name}` }
}
