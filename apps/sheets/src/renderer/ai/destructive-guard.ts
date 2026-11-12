/**
 * F16 (R3 triage): the F11 destructive gate keyed on the model's self-reported
 * `confirmDestructive` flag was bypassed by the 27B model itself —
 * - R3 T35: a destructive batch (delete_sheet) submitted ALREADY FLAGGED as
 *   the first submission, before any refusal and before the model asked the
 *   user — applied directly.
 * - R3 T8: a batch refused without the flag was resubmitted FLAGGED in the
 *   SAME model turn, with no user message in between — applied directly.
 *
 * The flag can therefore never be the release condition on its own. The gate
 * adds a mechanical, semantics-free requirement: the batch must have been
 * refused in an EARLIER USER RUN (a run = one `loop.run()` = one user
 * message). Only then has the user actually had a chance to see the refusal
 * and confirm. Same-run flagged retries and flagged-first submissions are
 * refused (and re-recorded), so every path converges on "refused once in a
 * previous run, then confirmed" — the worst case is one extra confirmation
 * round trip, never an unconfirmed destructive apply.
 *
 * State is session-scoped: one `DestructiveGateState` per
 * `createWorkbookSkill()` instance (see workbook-skill.ts), ticked by the
 * loop's per-user-run `onRunStart`.
 */
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

export interface DestructiveGateState {
  /** number of user runs started so far; ticks once per loop.run */
  runSeq: number
  /** stable batch key → runSeq of its last refusal */
  refusedAt: Map<string, number>
}

export function createDestructiveGateState(): DestructiveGateState {
  return { runSeq: 0, refusedAt: new Map() }
}

/** Bounded growth: keep at most this many pending refusals (a session can span hundreds of runs). */
const MAX_PENDING = 100

/**
 * Content-addressed key for a batch. Keys are recursively sorted so two
 * submissions of the same operations with different property orders still
 * collide (the model does not preserve key order across regenerations).
 */
export function stableBatchKey(operations: unknown): string {
  return stableStringify(operations)
}

function stableStringify(value: unknown): string {
  if (value === null || typeof value !== 'object') return JSON.stringify(value) ?? 'null'
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(',')}]`
  const record = value as Record<string, unknown>
  const keys = Object.keys(record)
    .filter((k) => record[k] !== undefined)
    .sort()
  return `{${keys.map((k) => `${JSON.stringify(k)}:${stableStringify(record[k])}`).join(',')}}`
}

export type DestructiveGateDecision =
  | { kind: 'apply' }
  | {
      kind: 'refuse'
      /** which refusal path produced this message (test/telemetry hook) */
      reason: 'not-flagged' | 'flagged-never-refused' | 'flagged-same-run'
      message: string
    }

/**
 * Decide a destructive batch against the gate state.
 *
 * - no flag → refuse, record the refusal at the current run (F11 behavior).
 * - flag, and the batch was refused in an EARLIER user run → apply, consume
 *   the record (the user had a turn to confirm between refusal and resubmit).
 * - flag, never refused → refuse + record: the model set the flag preemptively.
 * - flag, refused only within this run → refuse + re-record: no user turn has
 *   elapsed since the refusal, so no confirmation could have happened.
 */
export function destructiveGateCheck(args: {
  operations: readonly WorkbookOperation[]
  flagged: boolean
  /** the human-readable scale notes (from destructiveOperationNotes) */
  scale: readonly string[]
  state: DestructiveGateState
}): DestructiveGateDecision {
  const { operations, flagged, scale, state } = args
  const scaleText = scale.join('; ')
  const base = `Stopped — this batch is destructive at scale: ${scaleText}. Nothing was changed. `

  if (!flagged) {
    state.refusedAt.set(stableBatchKey(operations), state.runSeq)
    prune(state)
    return {
      kind: 'refuse',
      reason: 'not-flagged',
      message:
        base +
        `Tell the user what would be removed and ask for their explicit confirmation; ` +
        `only after they confirm, resubmit the SAME batch with confirmDestructive: true.`,
    }
  }

  const key = stableBatchKey(operations)
  const refusedAt = state.refusedAt.get(key)
  if (refusedAt !== undefined && refusedAt < state.runSeq) {
    // refused in an earlier user run: the user has had a turn to respond.
    state.refusedAt.delete(key)
    return { kind: 'apply' }
  }

  state.refusedAt.set(key, state.runSeq)
  prune(state)
  if (refusedAt === undefined) {
    return {
      kind: 'refuse',
      reason: 'flagged-never-refused',
      message:
        base +
        `It was submitted with confirmDestructive: true already set, before it was ever refused and the user asked — ` +
        `the flag is only honored after a refusal in an earlier turn. ` +
        `Submit it now WITHOUT the flag (the refusal will be recorded), tell the user exactly what would be removed, ` +
        `and only after they explicitly confirm, resubmit the SAME batch with confirmDestructive: true.`,
    }
  }
  return {
    kind: 'refuse',
    reason: 'flagged-same-run',
    message:
      base +
      `It was already refused earlier in this same turn, before the user could respond — ` +
      `confirmDestructive: true is only honored after the refusal AND a new user turn. ` +
      `Ask the user for their explicit confirmation now, and resubmit the SAME batch with the flag in their next reply.`,
  }
}

function prune(state: DestructiveGateState): void {
  if (state.refusedAt.size <= MAX_PENDING) return
  // Map preserves insertion order; drop the oldest entries.
  for (const key of state.refusedAt.keys()) {
    if (state.refusedAt.size <= MAX_PENDING) break
    state.refusedAt.delete(key)
  }
}
