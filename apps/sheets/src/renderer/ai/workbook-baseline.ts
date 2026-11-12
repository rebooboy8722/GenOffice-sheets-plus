/**
 * Freshness tracking between the AI's view of the workbook and the workbook
 * itself.
 *
 * The model reads a range, then proposes writes against the coordinates it
 * saw. If the user edits the grid in between, those coordinates may no longer
 * mean what the model believed — and the apply path only compares the planned
 * cells' *values* (planStillMatches in ../lazy-plan.ts), so a user who
 * inserted a row silently retargets every address and the write lands on the
 * wrong record with no error. Docs and Markdown guard this with a doc-identity
 * check (markDocSeen / editedExternally / STALE_DOC_ERROR in
 * apps/docs/src/renderer/ai/tools.ts); a sheet has no persistent object to
 * compare identity against, so this module tracks monotonic counters instead.
 *
 * Two counters, because a sheet addresses cells by A1 notation and not every
 * edit invalidates those coordinates:
 *
 *  - **structural** — the coordinate space itself moved, or a range stopped
 *    meaning what it did (insert/delete rows and columns, sheet
 *    add/remove/rename/reorder, merge/unmerge, reorder/move). Every address the
 *    model holds is now suspect, so writing to a range is refused until it
 *    re-reads.
 *  - **content** — values, formatting, sizes, filters, notes. `B5` still points
 *    at the same cell, so writes are not refused; the drift only matters as
 *    information ("what you read may no longer be what is there").
 *
 * The counters advance on **user** edits only. Programmatic writes — AI applies
 * and streaming file loads alike — raise a gate around themselves, and a
 * mutation seen under either gate is dropped. Without that filter the AI's own
 * write would bump the baseline and the next write in the same conversation
 * would then be blocked by its own previous turn.
 */

import {
  AXIS_ATTR_MUTATIONS,
  CF_MUTATIONS,
  DEFINED_NAME_MUTATIONS,
  DV_MUTATIONS,
  FILTER_MUTATIONS,
  MERGE_MUTATIONS,
  MOVE_RANGE_MUTATION,
  MOVE_ROWS_MUTATION,
  NOTE_MUTATIONS,
  REMOVE_NUMFMT_MUTATION,
  REORDER_RANGE_MUTATION,
  ROW_COLUMN_MUTATIONS,
  SET_NUMFMT_MUTATION,
  SET_RANGE_VALUES_MUTATION,
  SHEET_LIFECYCLE_MUTATIONS,
} from '../app-constants'
import { aiApplyGate, journalSuppression } from '../univer-state'
import type { ToolExecution } from './tools'

/** Which kind of drift a mutation records, or null when the AI need not care. */
export type DriftKind = 'structural' | 'content'

/**
 * Edits that leave every A1 address pointing at the same cell. Sourced from the
 * same constants the edit journal keys on (../app-constants.ts) rather than
 * guessed from id prefixes: Univer's own naming is not uniform
 * (`set.numfmt` is dotted where `set-range-values` is hyphenated), and the
 * view-state mutations that also start with `set-` (`set-frozen`,
 * `set-worksheet-right-to-left`) change nothing the model addresses.
 */
const CONTENT_MUTATIONS = new Set<string>([
  SET_RANGE_VALUES_MUTATION,
  SET_NUMFMT_MUTATION,
  REMOVE_NUMFMT_MUTATION,
  ...Object.keys(AXIS_ATTR_MUTATIONS),
  ...FILTER_MUTATIONS,
  ...CF_MUTATIONS,
  ...DV_MUTATIONS,
  ...DEFINED_NAME_MUTATIONS,
  ...NOTE_MUTATIONS,
])

export function driftKindOf(eventId: string): DriftKind | null {
  if (ROW_COLUMN_MUTATIONS[eventId] || MERGE_MUTATIONS[eventId]) return 'structural'
  if (
    SHEET_LIFECYCLE_MUTATIONS.has(eventId) ||
    eventId === REORDER_RANGE_MUTATION ||
    eventId === MOVE_RANGE_MUTATION ||
    eventId === MOVE_ROWS_MUTATION
  ) {
    return 'structural'
  }
  if (CONTENT_MUTATIONS.has(eventId)) return 'content'
  return null
}

/**
 * True while a write is programmatic rather than user-driven.
 *
 * Both gates are needed and neither alone is enough:
 *  - `journalSuppression` covers streaming loads, demo-mode applies, fill_range
 *    undo/redo and large-filter rebuilds.
 *  - `aiApplyGate` covers the lazy-mode AI apply, whose cell/format write loop
 *    deliberately stays *inside* the edit journal so ⌘Z and the save patch
 *    still see it — so journalSuppression never wraps it.
 *
 * It is **not** `aiBulkUndoGate`, which an e2e run caught making the whole gate
 * inert: ribbon and dialog edits run the same `applyChangePlan` executor, so
 * that gate is raised while the *user* clicks "Insert row" — suppressing the
 * exact edit this module exists to notice. `aiApplyGate` is raised only by the
 * AI's own call (`ApplyPlanOptions.aiInitiated`).
 */
function underProgrammaticGate(): boolean {
  return journalSuppression.active || aiApplyGate.active
}

const counters = { structural: 0, content: 0 }

let seen: { structural: number; content: number } | null = null

/**
 * Records a user edit against the model's baseline. Called from the shared
 * CommandExecuted listener, which every mutation reaches; the gates and the
 * formula-engine filter decide whether it counts as a user edit.
 */
export function noteWorkbookMutation(eventId: string, fromFormula: boolean): void {
  if (fromFormula) return
  const kind = driftKindOf(eventId)
  if (!kind) return
  if (underProgrammaticGate()) return
  counters[kind] += 1
}

/** The model has just been shown the current workbook; that is the baseline. */
export function markWorkbookSeen(): void {
  seen = { structural: counters.structural, content: counters.content }
}

/**
 * Identity of the workbook the counters describe — its on-disk path, supplied by
 * the caller. Not the Univer unit id or the sidecar session id: both change when
 * the user saves the same file, and a save must not read as a switch.
 *
 * The counters are module state and the conversation outlives a file: the shell
 * reuses a renderer when you open another workbook (`openLazyWorkbook` resets the
 * preview, revision, pending-edit count and slicers, but not the chat, and
 * `AgentLoop.reset()` runs only when the user clicks "New chat"). So the model can
 * be looking at a brand-new overview while its history still holds ranges read
 * from a **different file** — and because `buildContext` re-marks the baseline on
 * every turn, drift is silent there: nothing was edited, the workbook was
 * replaced. Docs cannot hit this because its guard compares document identity
 * (`markDocSeen` / `editedExternally`); a sheet needed the same notion.
 */
let identity: string | null = null

/**
 * Armed by `noteWorkbookDocumentReplaced`, consumed by the next
 * `noteWorkbookIdentity` call. Needed because a path comparison cannot see every
 * replacement: a brand-new unsaved workbook (or a converted import) has no
 * identity to compare, so opening a file over one of those would go unreported.
 */
let switchPending = false

function dropWorkbookHistory(): void {
  counters.structural = 0
  counters.content = 0
  seen = null
}

/**
 * Records which workbook is open and, when it is a different one, throws away the
 * previous workbook's edit history.
 *
 * Returns true when the workbook changed under an existing conversation — the
 * caller puts `WORKBOOK_SWITCHED_NOTE` in front of the context so the model learns
 * it was, rather than discovering it by writing into the wrong file. False on the
 * first call (nothing to compare against) and when the identity is unknown.
 *
 * Clearing the counters keeps them describing exactly one workbook. It is
 * hygiene, not the fix: in the current wiring `buildContext` re-marks the
 * baseline in the same turn (`executeWorkbookTool` has no other caller), so the
 * old counts would already have been rebased before any gate read them.
 */
export function noteWorkbookIdentity(next: string | null): boolean {
  if (switchPending) {
    // Reported once, on the turn after the load — the load itself has no context
    // to attach to. The counters go with the document that produced them.
    switchPending = false
    dropWorkbookHistory()
    if (next !== null) identity = next
    return true
  }
  if (next === null) return false
  const changed = identity !== null && identity !== next
  if (changed) dropWorkbookHistory()
  identity = next
  return changed
}

/**
 * Says "the document in this window was replaced", from the place that knows it:
 * the loader, when the load is not a save's session swap and a workbook was
 * already open.
 *
 * This is the primary signal; the path comparison in `noteWorkbookIdentity` is
 * the backstop for replacements that reach the renderer some other way. The two
 * overlap on a plain File > Open of another path, which is harmless — the flag is
 * consumed by the same turn that would have compared the paths.
 */
export function noteWorkbookDocumentReplaced(): void {
  identity = null
  switchPending = true
}

/**
 * Records a new identity **without** reporting a switch: the workbook was
 * reopened under a different path but it is the same document, so the
 * conversation and the edit counters still describe it.
 *
 * The one caller is a save that reopens the file (`openLazyWorkbook` with
 * `continueChat`, whose own comment reads "the reopen is a session swap over the
 * same document"). **Save As** takes that path with a different path, so without
 * this the model would be told "a different workbook is open, the ranges you
 * read may not exist" right after the user pressed Save As on the file it is
 * editing — a false alarm that costs a re-read and teaches the model to ignore
 * the note. Counters are deliberately kept: same document, same edits.
 */
export function realignWorkbookIdentity(next: string | null): void {
  if (next !== null) identity = next
}

/**
 * Context header for the case no counter can see: the file changed, so every
 * address in the conversation's history refers to some other workbook.
 */
export const WORKBOOK_SWITCHED_NOTE =
  '⚠️ A DIFFERENT WORKBOOK IS NOW OPEN than earlier in this conversation (the user opened ' +
  'another file in this window). Any range, sheet name, or value you read earlier belongs to ' +
  'that other file and may not exist here. Do NOT reuse those addresses. The overview below is ' +
  'the only current truth — call read_range before writing, and tell the user which workbook ' +
  'you are editing before you change it.'

export interface WorkbookDrift {
  /** addresses the model holds may no longer point where it believes */
  structural: boolean
  /** values/formatting changed under the model; addresses still valid */
  content: boolean
}

/** Drift since the model was last shown the workbook. */
export function workbookDrift(): WorkbookDrift {
  if (!seen) return { structural: false, content: false }
  return {
    structural: counters.structural !== seen.structural,
    content: counters.content !== seen.content,
  }
}

/**
 * Runs a wait that cannot mutate the workbook with the AI-authored gate lowered,
 * and reports whether the coordinate space moved while it ran.
 *
 * The gate exists to tell the AI's own mutations from the user's, so it must
 * cover every step that *writes* — but it has no business covering a wait that
 * writes nothing. `pinStreamedPrecedents` polls the sidecar for up to 60 s with
 * the grid fully responsive, and a row the user inserted in that window was
 * dropped as "AI-authored"; the batch then wrote at coordinates the user had
 * already moved — the exact silent miswrite this module exists to prevent.
 *
 * Why lowering the flag cannot let the pin's own work count as a user edit:
 * the pin does install cells, through `patchWorksheetRange`, which raises
 * `journalSuppression` around every write it makes (univer-sync.ts). Since
 * `underProgrammaticGate` is a **union** of the two flags, that work stays
 * excluded with this one down. The rule for a new call site is therefore:
 * lower the gate only around work whose writes are self-suppressed, or that
 * writes nothing at all.
 *
 * Lowering the gate only makes such an edit *visible*, so the caller has to act
 * on the returned drift — `STALE_MID_APPLY_ERROR` below is what an apply should
 * throw. Only drift that appeared **during the window** is reported: drift
 * predating it is the caller's pre-existing business (the gate at
 * `gateWorkbookWrite` already saw it), not this wait's.
 *
 * What lowering it means, precisely: "nothing between here and there is ours".
 * It changes no journal, undo-batch or save behaviour — the mutations still
 * commit and stay undoable exactly as before. `underProgrammaticGate` below is
 * the flag's only reader (`grep -rn aiApplyGate apps packages`), so this is the
 * flag's whole second job: for the duration of the wait, a ⌘Z or an inserted row
 * counts as the user's again. The sibling `journalSuppression` is what hides
 * writes from the save journal, and it is untouched here — deliberately, since
 * lowering *that* one would stop the AI's writes being saved at all.
 */
export async function outsideAiGate<T>(
  aiInitiated: boolean,
  wait: () => Promise<T>,
): Promise<{ value: T; drift: WorkbookDrift }> {
  const none: WorkbookDrift = { structural: false, content: false }
  if (!aiInitiated) return { value: await wait(), drift: none }
  const before = workbookDrift()
  const wasAuthored = aiApplyGate.active
  aiApplyGate.active = false
  try {
    const value = await wait()
    const after = workbookDrift()
    return {
      value,
      drift: {
        structural: after.structural && !before.structural,
        content: after.content && !before.content,
      },
    }
  } finally {
    // Restored even when the wait fails: everything after it still writes as
    // the AI's own work. Restored to what the caller had, not forced on — a
    // caller that never claimed authorship must not end up claiming it.
    aiApplyGate.active = wasAuthored
  }
}

/**
 * Returned in place of executing a range-addressed write whose coordinates may
 * have moved. Worded as an instruction, like the docs equivalent: a model that
 * reads it knows the call to make next.
 */
export const STALE_WORKBOOK_ERROR =
  'The workbook structure changed since you last read it (rows/columns or sheets were ' +
  'inserted, removed, moved, or merged), so the addresses in your plan may no longer point ' +
  'where you intend. Do NOT submit these operations. Call get_workbook_context (and ' +
  'read_range for the region you are writing) to get the current layout, then rebuild the ' +
  'operations against the fresh coordinates.'

/**
 * The same event one step later: the coordinate space moved **while** a batch
 * was applying, so nothing further is written.
 *
 * Deliberately conditional about how much landed, because this function cannot
 * know: an apply throws from here either before its first operation (the pin
 * runs ahead of the cell loop) or after several have committed. The caller
 * reports which — `ApplyOutcome.partiallyApplied`, worded by `propose_operations`
 * as either "MID-BATCH … already committed" or "the workbook is UNCHANGED" — so
 * asserting either state here would contradict it, and a model handed "nothing
 * changed" plus "part of it landed" trusts neither.
 */
export const STALE_MID_APPLY_ERROR =
  'The workbook structure changed while this batch was applying (rows/columns or sheets were ' +
  'inserted, removed, moved, or merged partway through), so the rest of the batch was aborted. ' +
  'Call get_workbook_context and read_range before continuing, and rebuild anything still ' +
  'missing against the fresh coordinates — if operations from this batch already landed, do NOT ' +
  're-run those, or they will apply twice.'

/**
 * Content drift discovered mid-apply, reported through `ApplyOutcome.notices`
 * rather than refused: the addresses still point at the cells the plan named.
 * Phrased as a notice fragment (joined after "Note:" in ai/tools.ts).
 */
export const MID_APPLY_CONTENT_NOTICE =
  'the user edited cell values while this batch was applying — the addresses still match, ' +
  'so the change above stands, but re-read before reasoning from values you quoted earlier'

/**
 * Tools that read the workbook back into context. Running one means the model
 * now holds the current state, so the baseline moves with it — the same role
 * `markDocSeen` plays after a docs read (apps/docs/.../tools.ts settle()).
 */
const REFRESH_TOOLS = new Set([
  'get_workbook_context',
  'read_range',
  'read_cells',
  'read_formats',
  'read_sheet_features',
  'find_cells',
  'aggregate_range',
  'trace_precedents',
  'trace_dependents',
])

/**
 * Refuse range-addressed writes whose coordinates may have moved, and move the
 * baseline whenever the model is handed fresh state.
 *
 * Reads are never refused (the model needs them to recover), and a write the AI
 * applied itself refreshes the baseline because the tool output already reports
 * what landed — including the read-back values.
 *
 * `mutated` is set only by propose_operations (tools.ts), which is the sole
 * write path in the workbook tool set.
 */
export function gateWorkbookWrite(
  toolName: string,
  fail: (summary: string, output: string) => ToolExecution,
  execute: () => ToolExecution | Promise<ToolExecution>,
): ToolExecution | Promise<ToolExecution> {
  const isWrite = toolName === 'propose_operations'
  const drift = workbookDrift()
  if (isWrite && drift.structural) {
    return fail('Workbook changed', STALE_WORKBOOK_ERROR)
  }
  const finish = (exec: ToolExecution): ToolExecution => {
    settleBaseline(toolName, exec)
    // Values the model wrote may have been edited under it since it read them.
    // Addresses held, so the write stands — but a model that believes it is
    // correcting a column it read five turns ago should know the ground moved.
    // Only on a write: attaching it to reads would bury the read output itself.
    if (isWrite && drift.content && !exec.isError) {
      return { ...exec, output: `${exec.output}\n${CONTENT_DRIFT_NOTE}` }
    }
    return exec
  }
  const result = execute()
  return isPromiseLike(result) ? Promise.resolve(result).then(finish) : finish(result)
}

/**
 * Content drift is reported, never enforced: the addresses still point where
 * they did, so refusing the write would trade a real edit for a round trip
 * that re-confirms a coordinate nobody moved.
 */
const CONTENT_DRIFT_NOTE =
  'Note: the workbook contents changed since you last read them (the user edited cells or ' +
  'formatting between your read and this write). Addresses still point at the same cells, so ' +
  'the change above stands — but anything you quoted from an earlier read may no longer be ' +
  'current. Re-read before reasoning from old values.'

function settleBaseline(toolName: string, exec: { mutated: boolean; isError?: boolean }): void {
  if (exec.isError) return
  if (exec.mutated || REFRESH_TOOLS.has(toolName)) markWorkbookSeen()
}

function isPromiseLike<T>(value: T | Promise<T>): value is Promise<T> {
  return typeof (value as Promise<T>)?.then === 'function'
}
