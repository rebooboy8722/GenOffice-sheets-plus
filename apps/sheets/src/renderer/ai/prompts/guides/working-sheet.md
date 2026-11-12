# Working sheet for multi-step computation (working-sheet)

Do the hard arithmetic **on a scratch working sheet**, not in the user's real
sheet. The real sheet receives only the **final, verified results**; every
intermediate formula and every trial-and-error attempt lives in the working
sheet, which is removed when the task is done. This keeps the user's data clean
and means that if the run is interrupted, their sheet still holds the original
data plus clean final answers — not a pile of helper formulas.

## When to use it

Use a working sheet when **any** of these is true:
- The task is a **multi-step or non-trivial computation** — aggregations over a
  large table, cross-column conditions, or a value derived from another computed
  value.
- You are **not certain which formula will work** and will need to try variants
  and check the results before committing to an answer.
- The result of one step feeds the next (compute A, then compute B from A).

For a single formula you are confident in (one `SUM`, one `COUNTIFS`), skip the
working sheet: use `evaluate_formula` (no cell write) or write the formula
directly.

## The workflow

1. **Create the working sheet** — structural, so it goes in a batch **by itself**:
   `{op:"add_sheet", name:"GenOffice Scratch", rows:<size to the work>, columns:<…>}`.
   Size it with `rows`/`columns` so you don't write past its grid.
2. **Get its id** — call `get_workbook_context`; the id is generated, never guess it.
3. **Compute there** — a content batch. Write your candidate formulas into the
   working sheet. When you're unsure, **write several variants at once** and read
   back **all** the results in the one batch — do not test them one at a time in
   the user's real cells. `propose_operations` reads formula results back
   automatically, so the result tells you each candidate's value.
4. **Write the final results** — to the real target cells, **in the form the user
   asked for** (plain numbers, not formulas, when they asked for plain numbers),
   in one batch. This is the only time the real sheet is written for this task.
5. **Delete the working sheet** — structural, in a batch by itself:
   `{op:"delete_sheet", sheetId:<the working sheet's id>}`. Delete it yourself so
   you don't depend on the app's cleanup; the app removes any leftover working
   sheet when the run ends as a safety net.

## Rules

- **Never do exploratory or helper computation in the user's real sheet.** The
  real sheet must never hold a helper formula you then overwrite — write the
  final result there directly.
- A working sheet is a separate temporary **tab**, not part of the user's data;
  creating it does **not** count as changing the user's cells or adding/removing
  their rows or columns.
- Usually one working sheet is enough; if you need more space, use separate
  regions of the same sheet rather than creating more.
- Keep intermediate labels in the working sheet so you can read back which value
  is which (a header row naming each computed quantity).

## Formula gotchas (verified in this app)

These cost many wasted retries when you don't know them; test the safe variant
in the working sheet (or with `evaluate_formula`) before committing it to the
user's real cells:

- **Boolean arrays don't sum on their own.** `=SUMPRODUCT(VALUE(N2:N1258)>1000000)`
  returns **0**, not the count — the comparison produces booleans and `SUMPRODUCT`
  ignores bare booleans. Coerce them to numbers:
  `=SUMPRODUCT(--(VALUE(N2:N1258)>1000000))`, `=SUMPRODUCT((VALUE(N2:N1258)>1000000)*1)`,
  or `=SUMPRODUCT(IF(VALUE(N2:N1258)>1000000,1,0))`. Write the bare form and one
  coerced form side by side and read back — the coerced one gives the expected
  magnitude, the bare one gives 0.
- **Numbers stored as text** (the read path flags these as `stored as: TEXT`):
  wrap in `VALUE()` before any arithmetic or comparison, or they behave as 0/text
  and the answer comes out 0.
- **Whole-column formulas over a large (streaming) sheet undercount until the
  column is loaded.** A `COUNTIF`/`SUMPRODUCT` over `C2:C1258` only sees the cells
  currently materialized, so it can return a far-too-small number on the first
  try and the right number later. Before trusting a whole-column result, force the
  column to load (an `aggregate_range` or `read_range` over it), then re-check the
  formula. If two runs of the same whole-column formula give different counts,
  that is the cause — not the formula.
