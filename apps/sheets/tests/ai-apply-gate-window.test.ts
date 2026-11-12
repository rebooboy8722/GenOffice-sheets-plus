import { beforeEach, describe, expect, it, vi } from 'vitest'

import {
  MID_APPLY_CONTENT_NOTICE,
  markWorkbookSeen,
  noteWorkbookMutation,
  workbookDrift,
} from '../src/renderer/ai/workbook-baseline'
import { createEditJournal } from '../src/renderer/edit-journal'
import { applyChangePlan } from '../src/renderer/op-executor'
import type { ApplyOutcome, ChangePlan } from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'
import { aiApplyGate, journalSuppression } from '../src/renderer/univer-state'

/**
 * The AI-authored gate inside an apply, and the user edit it must not swallow.
 *
 * `aiApplyGate` is what tells the freshness baseline (ai/workbook-baseline.ts)
 * "this mutation came from the AI, not from the user". It is raised for the whole
 * apply, but an apply also *waits* — pinning streamed formula precedents polls
 * the sidecar for up to 60 s — and the grid stays clickable the entire time. A
 * row the user inserted during such a wait used to be filed as the AI's own and
 * dropped, after which the batch wrote at coordinates the user had already moved:
 * the silent miswrite the baseline exists to prevent, arriving through a hole in
 * the gate rather than through a missing check.
 *
 * So the wait has to run with the gate lowered (`outsideAiGate`) *and* the apply
 * has to act on what that reveals. Both halves are tested here against the real
 * `applyChangePlan`, with only the sidecar pin faked — the fake is where the user
 * edit lands, which is exactly the race.
 */

const INSERT_ROW = 'sheet.mutation.insert-row'
const SET_VALUES = 'sheet.mutation.set-range-values'

/** Injected into the window: what the user does while the apply is parked. */
let duringPin: (() => void) | null = null
/** Injected during the prechecks, which run before the gate is raised. */
let duringPrecheck: (() => void) | null = null

/** What the pin could observe, i.e. what the apply was claiming about itself. */
const pinSaw: { calls: number; gateActive: boolean | null } = {
  calls: 0,
  gateActive: null,
}

vi.mock('../src/renderer/univer-sync', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/renderer/univer-sync')>()
  return {
    ...real,
    // File-state capture is a sidecar poll of its own, and it runs in the
    // prechecks *above* the gate — the precheck-phase race is injected here.
    ensureSheetFileState: async () => {
      duringPrecheck?.()
      return true
    },
    pinStreamedPrecedents: async () => {
      pinSaw.calls += 1
      pinSaw.gateActive = aiApplyGate.active
      duringPin?.()
      // A real poll yields for milliseconds to seconds; the await is what lets
      // an edit land "inside" the wait.
      await new Promise((resolve) => setTimeout(resolve, 0))
      return true
    },
  }
})

const written: string[] = []

function fakeWorkbook() {
  const sheet = {
    getSheetId: () => 'sh1',
    getSheetName: () => 'Sheet1',
    // Enough of a structural op for the "an op already landed" case: the point
    // is that it commits *before* the pin, so the abort is genuinely mid-batch.
    insertRowsBefore: (row: number, count: number) => {
      written.push(`insert@${row}+${count}`)
      noteWorkbookMutation(SET_VALUES, false)
    },
    getRange: (address: string) => ({
      setFormula: (formula: string) => {
        written.push(`${address}=${formula}`)
        // The real write reaches the app's CommandExecuted listener, which is
        // where the baseline reads from: an AI write must not count as drift.
        noteWorkbookMutation(SET_VALUES, false)
      },
      setValues: () => {
        written.push(address)
        noteWorkbookMutation(SET_VALUES, false)
      },
      getCellDatas: () => [[{}]],
      clearContent: () => {},
    }),
  }
  return {
    getId: () => 'unit-1',
    getSheets: () => [sheet],
    getSheetBySheetId: (id: string) => (id === 'sh1' ? sheet : undefined),
    getActiveSheet: () => sheet,
    getWorkbook: () => ({ getStyles: () => [] }),
  }
}

function fakeRuntime(workbook: ReturnType<typeof fakeWorkbook>): UniverRuntime {
  return {
    univerAPI: { getActiveWorkbook: () => workbook },
    univer: {
      __getInjector: () => ({
        get: () => ({ __tempBatchingUndoRedo: () => ({ dispose: () => {} }) }),
      }),
    },
  } as unknown as UniverRuntime
}

function streamedState(): LazyWorkbookState {
  return {
    file: {
      sessionId: 'session-1',
      name: 'book.xlsx',
      // A streamed (not fully loaded) workbook is what puts the pin on this path.
      // The per-sheet shape is what the real gates read (lazyGateFailure looks at
      // pivotRanges / sourceXmlBytes), so it is spelled out rather than stubbed.
      sheets: [
        {
          id: 'sh1',
          name: 'Sheet1',
          rowCount: 100,
          columnCount: 10,
          sourceXmlBytes: 0,
          pivotRanges: [],
          tables: [],
          pivotTables: [],
          merges: [],
        },
      ],
      styles: [],
      visuals: [],
    },
    generation: 1,
    loadedRanges: new Map(),
    loadingKeys: new Map(),
    retryTimers: new Map(),
    appliedMerges: new Map(),
    appliedRowKeys: new Map(),
    sheetProtections: new Map(),
    sheetPageBreaks: new Map(),
    sheetProtectedRanges: new Map(),
    uninstalledDefinedNames: new Set(),
    appliedCfSheets: new Set(),
    appliedFilterSheets: new Set(),
    appliedDvSheets: new Set(),
    decorationsPendingSheets: new Set(),
    hyperlinkTargets: new Map(),
    frozenStripKeys: new Map(),
    filterOrigins: new Map(),
    showFormulaSheets: new Set(),
    formulaMode: false,
    editJournal: createEditJournal(),
    flags: { preloadComplete: false },
    closure: { status: 'idle', pinned: new Map() },
    formulaText: new Map(),
    cachedFormulaValues: new Map(),
    pivotDefinitions: new Map(),
    outline: new Map(),
    recalc: { timer: null, generation: 0, failures: 0, formulaCells: new Map(), overlay: new Map() },
  } as unknown as LazyWorkbookState
}

/** One formula write: the formula is what makes the apply pin precedents. */
const plan: ChangePlan = {
  transactionId: 'tx-1',
  baseRevision: 0,
  cellChanges: [
    {
      sheetId: 'sh1',
      address: 'A10',
      before: { value: null },
      after: { value: null, formula: '=B2*2' },
    },
  ],
  sheetRenames: [],
  structuralChanges: [],
  formatChanges: [],
  warnings: [],
}

async function apply(options: { aiInitiated?: boolean } = {}): Promise<ApplyOutcome> {
  const state = streamedState()
  const workbook = fakeWorkbook()
  return applyChangePlan(plan, {
    runtime: fakeRuntime(workbook),
    workbook: workbook as never,
    state,
    lazyWorkbookRef: { current: state },
    setMessage: () => {},
    setPendingEdits: () => {},
  } as never, options)
}

describe('the AI-authored gate inside an apply', () => {
  beforeEach(() => {
    aiApplyGate.active = false
    journalSuppression.active = false
    duringPin = null
    duringPrecheck = null
    written.length = 0
    pinSaw.calls = 0
    pinSaw.gateActive = null
    markWorkbookSeen()
  })

  it('does not hold the gate across a wait that cannot mutate the grid', async () => {
    const outcome = await apply({ aiInitiated: true })
    expect(outcome.ok).toBe(true)
    // The pin is mutation-free sidecar IO: inside it, the apply must not be
    // claiming authorship of whatever the user does.
    expect(pinSaw.calls).toBe(1)
    expect(pinSaw.gateActive).toBe(false)
  })

  it('aborts the batch when the user moves the coordinate space during that wait', async () => {
    duringPin = () => noteWorkbookMutation(INSERT_ROW, false)
    const outcome = await apply({ aiInitiated: true })
    expect(outcome.ok).toBe(false)
    expect(outcome.reason).toContain('while this batch was applying')
    // Nothing was written: the refusal happens before the cell loop, so the
    // addresses the plan holds are never used against the shifted grid.
    expect(written).toEqual([])
    expect(outcome.partiallyApplied).toBeUndefined()
  })

  it('does not claim a half-applied batch when nothing had landed', async () => {
    duringPin = () => noteWorkbookMutation(INSERT_ROW, false)
    const outcome = await apply({ aiInitiated: true })
    // The pin runs AHEAD of the cell loop, so this abort leaves no operation
    // committed and propose_operations words the outcome "the workbook is
    // UNCHANGED". An error text asserting part of it landed would contradict
    // the sentence wrapped around it, and a model handed both trusts neither.
    expect(outcome.partiallyApplied).toBeUndefined()
    expect(outcome.reason).not.toMatch(/some of the batch landed|were already committed/)
    // Actionable in both directions: re-read, and never re-apply what landed.
    expect(outcome.reason).toContain('do NOT re-run those')
  })

  it('reports the same abort as partial once an operation has landed', async () => {
    // The other half of the wording contract: with an op committed, the outcome
    // must say so — the shared text stays silent either way and lets this flag
    // carry the fact.
    duringPin = () => noteWorkbookMutation(INSERT_ROW, false)
    const state = streamedState()
    const workbook = fakeWorkbook()
    const outcome = await applyChangePlan(
      { ...plan, structuralChanges: [{ op: { op: 'insert_rows', sheetId: 'sh1', row: 2, count: 1 }, label: 'insert a row' }] },
      {
        runtime: fakeRuntime(workbook),
        workbook: workbook as never,
        state,
        lazyWorkbookRef: { current: state },
        setMessage: () => {},
        setPendingEdits: () => {},
      } as never,
      { aiInitiated: true },
    )
    expect(outcome.ok).toBe(false)
    // Named explicitly: without this, any unrelated throw inside the apply
    // would satisfy "ok === false" and the test would pass for the wrong reason.
    expect(outcome.reason).toContain('while this batch was applying')
    expect(outcome.partiallyApplied).toBe(true)
    // The row insert committed; the formula write that would have used the
    // stale coordinates did not.
    expect(written).toEqual(['insert@1+1'])
  })

  it('leaves the drift unresolved so the next write is refused too', async () => {
    duringPin = () => noteWorkbookMutation(INSERT_ROW, false)
    await apply({ aiInitiated: true })
    // Clearing it here would let the batch's own refusal count as a re-read.
    expect(workbookDrift().structural).toBe(true)
  })

  it('lets a content edit through, reported as a notice', async () => {
    duringPin = () => noteWorkbookMutation(SET_VALUES, false)
    const outcome = await apply({ aiInitiated: true })
    // B10 still points at the same cell, so refusing would trade a real edit for
    // a round trip that re-confirms a coordinate nobody moved.
    expect(outcome.ok).toBe(true)
    expect(outcome.notices).toContain(MID_APPLY_CONTENT_NOTICE)
    expect(written).toEqual(['A10==B2*2'])
  })

  it('still treats the AI’s own writes as authorship after the wait', async () => {
    // The gate must come back up: without the restore, every cell write below
    // the pin would count as a user edit and the apply would block its own next
    // turn.
    const outcome = await apply({ aiInitiated: true })
    expect(outcome.ok).toBe(true)
    expect(written.length).toBe(1)
    expect(workbookDrift()).toEqual({ structural: false, content: false })
    expect(aiApplyGate.active).toBe(false)
  })

  it('refuses before any write when the user restructures during the prechecks', async () => {
    duringPrecheck = () => noteWorkbookMutation(INSERT_ROW, false)
    const outcome = await apply({ aiInitiated: true })
    // The prechecks run with the gate down, so this edit is already counted;
    // verifyBeforeApply only compares cell values (vacuous for bulk ops), so a
    // dedicated structural check is what stops the batch here.
    expect(outcome.ok).toBe(false)
    expect(outcome.reason).toContain('structure changed since you last read')
    expect(written).toEqual([])
    // Nothing had written → the tool reports "the workbook is UNCHANGED".
    expect(outcome.partiallyApplied).toBeUndefined()
  })

  it('lets a content edit during the prechecks through', async () => {
    duringPrecheck = () => noteWorkbookMutation(SET_VALUES, false)
    const outcome = await apply({ aiInitiated: true })
    expect(outcome.ok).toBe(true)
    expect(written).toEqual(['A10==B2*2'])
  })

  it('leaves a ribbon edit alone: no gate, no abort', async () => {
    duringPin = () => noteWorkbookMutation(INSERT_ROW, false)
    const outcome = await apply({})
    // The UI path never raises the gate, so there is nothing to lower and no
    // baseline semantics to enforce — this must behave exactly as before.
    expect(pinSaw.gateActive).toBe(false)
    expect(outcome.ok).toBe(true)
    expect(written).toEqual(['A10==B2*2'])
    // The edit still counts as the user's own for whatever the AI does next.
    expect(workbookDrift().structural).toBe(true)
  })
})
