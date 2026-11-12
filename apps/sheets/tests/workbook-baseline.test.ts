import { beforeEach, describe, expect, it } from 'vitest'

import {
  driftKindOf,
  gateWorkbookWrite,
  markWorkbookSeen,
  noteWorkbookDocumentReplaced,
  noteWorkbookIdentity,
  noteWorkbookMutation,
  outsideAiGate,
  realignWorkbookIdentity,
  workbookDrift,
} from '../src/renderer/ai/workbook-baseline'
import { aiApplyGate, aiBulkUndoGate, journalSuppression } from '../src/renderer/univer-state'
import type { ToolExecution } from '../src/renderer/ai/tools'

const INSERT_ROW = 'sheet.mutation.insert-row'
const SET_VALUES = 'sheet.mutation.set-range-values'
const SET_NUMFMT = 'sheet.mutation.set.numfmt'
const SET_FROZEN = 'sheet.mutation.set-frozen'
const REMOVE_SHEET = 'sheet.mutation.remove-sheet'

const wrote = (): ToolExecution => ({ output: 'applied', mutated: true, summary: 'applied' })
const read = (output = 'ctx'): ToolExecution => ({ output, mutated: false, summary: 'read' })
const failed = (output = 'boom'): ToolExecution => ({
  output,
  mutated: false,
  summary: 'failed',
  isError: true,
})

const fail = (summary: string, output: string): ToolExecution => ({
  output,
  isError: true,
  mutated: false,
  summary,
})

/**
 * Runs a tool through the gate and returns what the model would see. These
 * inputs resolve synchronously, so the Promise branch is not exercised here.
 */
function runGate(name: string, exec: ToolExecution): ToolExecution {
  return gateWorkbookWrite(name, fail, () => exec) as ToolExecution
}

/** propose_operations is the only write tool, so the only gated one. */
const propose = (exec = wrote()) => runGate('propose_operations', exec)
const readRange = (exec = read()) => runGate('read_range', exec)

describe('workbook freshness baseline', () => {
  beforeEach(() => {
    aiApplyGate.active = false
    aiBulkUndoGate.active = false
    journalSuppression.active = false
    // Baselines are relative, so re-marking clears any drift from a prior case.
    markWorkbookSeen()
  })

  describe('classification', () => {
    it('calls row/column and sheet changes structural', () => {
      expect(driftKindOf(INSERT_ROW)).toBe('structural')
      expect(driftKindOf(REMOVE_SHEET)).toBe('structural')
    })

    it('calls value and format edits content', () => {
      expect(driftKindOf(SET_VALUES)).toBe('content')
      // Univer names this one with a dot, not a hyphen.
      expect(driftKindOf(SET_NUMFMT)).toBe('content')
    })

    it('ignores view state the model never addresses', () => {
      expect(driftKindOf(SET_FROZEN)).toBeNull()
      noteWorkbookMutation(SET_FROZEN, false)
      expect(workbookDrift()).toEqual({ structural: false, content: false })
    })
  })

  describe('user edits reach the baseline', () => {
    it('a structural edit marks addresses stale', () => {
      noteWorkbookMutation(INSERT_ROW, false)
      expect(workbookDrift().structural).toBe(true)
    })

    it('a content edit does not', () => {
      noteWorkbookMutation(SET_VALUES, false)
      expect(workbookDrift()).toEqual({ structural: false, content: true })
    })

    it('ignores formula-engine derived writes', () => {
      noteWorkbookMutation(SET_VALUES, true)
      noteWorkbookMutation(INSERT_ROW, true)
      expect(workbookDrift()).toEqual({ structural: false, content: false })
    })
  })

  describe('programmatic writes must not count as user edits', () => {
    it('AI apply (aiApplyGate) is not drift', () => {
      // The lazy write loop stays inside the edit journal on purpose, so the
      // journal gate does NOT cover it — without this the AI would block itself.
      aiApplyGate.active = true
      noteWorkbookMutation(SET_VALUES, false)
      noteWorkbookMutation(INSERT_ROW, false)
      aiApplyGate.active = false
      expect(workbookDrift()).toEqual({ structural: false, content: false })
    })

    it('a ribbon edit is drift even though it raises the undo-batch gate', () => {
      // Ribbon "Insert row" runs the same applyChangePlan executor as an AI
      // plan, so aiBulkUndoGate is raised for a *user* edit too. Suppressing on
      // it made the whole gate inert (an e2e run caught it) — only aiApplyGate,
      // set solely by the AI's own call, means "the AI authored this".
      aiBulkUndoGate.active = true
      noteWorkbookMutation(INSERT_ROW, false)
      aiBulkUndoGate.active = false
      expect(workbookDrift().structural).toBe(true)
      expect(propose().isError).toBe(true)
    })

    it('streaming loads (journalSuppression) are not drift', () => {
      journalSuppression.active = true
      noteWorkbookMutation(SET_VALUES, false)
      noteWorkbookMutation(INSERT_ROW, false)
      journalSuppression.active = false
      expect(workbookDrift()).toEqual({ structural: false, content: false })
    })
  })

  describe('write gate', () => {
    it('refuses a write after a structural edit', () => {
      noteWorkbookMutation(INSERT_ROW, false)
      const result = propose()
      expect(result.isError).toBe(true)
      expect(result.output).toContain('structure changed')
    })

    it('does not execute the refused write', () => {
      noteWorkbookMutation(INSERT_ROW, false)
      let executed = false
      gateWorkbookWrite('propose_operations', fail, () => {
        executed = true
        return wrote()
      })
      expect(executed).toBe(false)
    })

    it('a refused write does not clear the drift', () => {
      // Clearing on refusal would let the model bypass the gate by retrying
      // until the counters happened to line up.
      noteWorkbookMutation(INSERT_ROW, false)
      propose()
      expect(workbookDrift().structural).toBe(true)
      expect(propose().isError).toBe(true)
    })

    it('allows a write after a content-only edit, and reports the drift', () => {
      noteWorkbookMutation(SET_VALUES, false)
      const result = propose()
      // The write stands (addresses still valid) but the model learns the ground
      // moved, so it stops reasoning from values it read earlier.
      expect(result.isError).toBeUndefined()
      expect(result.output).toContain('applied')
      expect(result.output).toContain('contents changed')
    })

    it('reports nothing when nothing moved', () => {
      expect(propose().output).toBe('applied')
    })

    it('reports no drift on a failed write', () => {
      noteWorkbookMutation(SET_VALUES, false)
      expect(propose(failed()).output).toBe('boom')
    })
  })

  describe('baseline settlement', () => {
    it('a re-read clears the drift', () => {
      noteWorkbookMutation(INSERT_ROW, false)
      readRange()
      expect(workbookDrift()).toEqual({ structural: false, content: false })
      expect(propose().isError).toBeUndefined()
    })

    it('a failed read does not clear the drift', () => {
      // Otherwise the model could clear the gate with a read that returned
      // nothing, then write against coordinates it never actually saw.
      noteWorkbookMutation(INSERT_ROW, false)
      readRange(failed('bad range'))
      expect(workbookDrift().structural).toBe(true)
    })

    it('a successful AI write clears the drift it is reporting', () => {
      noteWorkbookMutation(SET_VALUES, false)
      propose()
      expect(workbookDrift()).toEqual({ structural: false, content: false })
    })

    it('a failed write does not clear the drift', () => {
      noteWorkbookMutation(SET_VALUES, false)
      propose(failed())
      expect(workbookDrift().content).toBe(true)
    })

    it('reads are never refused, even while stale', () => {
      noteWorkbookMutation(INSERT_ROW, false)
      expect(readRange().isError).toBeUndefined()
    })
  })

  describe('outsideAiGate: a wait that cannot mutate the grid', () => {
    it('lowers the gate for the duration and restores it after', async () => {
      aiApplyGate.active = true
      const seen: boolean[] = []
      const { value, drift } = await outsideAiGate(true, async () => {
        seen.push(aiApplyGate.active)
        noteWorkbookMutation(INSERT_ROW, false)
        return 'done'
      })
      expect(value).toBe('done')
      expect(drift.structural).toBe(true)
      expect(seen).toEqual([false])
      expect(aiApplyGate.active).toBe(true)
      // The edit landed while the gate was down, so the apply must see it.
      expect(workbookDrift().structural).toBe(true)
    })

    it('restores the gate when the wait fails', async () => {
      // Everything after the wait still writes as the AI's own work, so a
      // rejection must not leave the gate down.
      aiApplyGate.active = true
      await expect(
        outsideAiGate(true, async () => {
          throw new Error('sidecar died')
        }),
      ).rejects.toThrow('sidecar died')
      expect(aiApplyGate.active).toBe(true)
    })

    it('reports only drift that appeared during the window', async () => {
      // Drift predating the wait is the caller's pre-existing business — the
      // tool-level gate already saw it, and reporting it again would blame an
      // innocent wait.
      noteWorkbookMutation(INSERT_ROW, false)
      const first = await outsideAiGate(true, async () => {})
      expect(first.drift).toEqual({ structural: false, content: false })
      // Restored to what the caller had, not forced on: a caller that never
      // claimed authorship must not end up claiming it.
      expect(aiApplyGate.active).toBe(false)
      markWorkbookSeen()
      const second = await outsideAiGate(true, async () => {
        noteWorkbookMutation(INSERT_ROW, false)
      })
      expect(second.drift.structural).toBe(true)
    })

    it('still excludes work the callee suppresses itself', async () => {
      // The safety property the apply window depends on: `pinStreamedPrecedents`
      // installs cells through patchWorksheetRange, which raises
      // journalSuppression around its own writes. With only this flag down, the
      // union in underProgrammaticGate must keep that work out of the baseline —
      // otherwise every formula write would report the AI's own pin as a user
      // edit and abort itself.
      aiApplyGate.active = true
      await outsideAiGate(true, async () => {
        journalSuppression.active = true
        try {
          noteWorkbookMutation(SET_VALUES, false)
          noteWorkbookMutation(INSERT_ROW, false)
        } finally {
          journalSuppression.active = false
        }
      })
      expect(workbookDrift()).toEqual({ structural: false, content: false })
    })

    it('does not touch the gate on a UI apply', async () => {
      // Ribbon edits never raise the gate; the helper must neither lower nor
      // raise it, or it would silently change what counts as authorship there.
      const seen: boolean[] = []
      const { value, drift } = await outsideAiGate(false, async () => {
        seen.push(aiApplyGate.active)
        noteWorkbookMutation(INSERT_ROW, false)
        return 1
      })
      expect(value).toBe(1)
      expect(seen).toEqual([false])
      expect(aiApplyGate.active).toBe(false)
      // No baseline work either: the UI path has no model to keep fresh.
      expect(drift).toEqual({ structural: false, content: false })
    })
  })

  describe('noteWorkbookIdentity: the file changed under the conversation', () => {
    // Identity is module state that outlives a test, so each case opens on its
    // own unique path and only asserts about transitions after that. Paths, not
    // unit ids: a save reopens the same file under a new unit, and that is not a
    // switch (an e2e run is what established this).
    const open = (path: string): void => {
      noteWorkbookIdentity(path)
      markWorkbookSeen()
    }

    it('says nothing while the same workbook stays open', () => {
      open('/books/budget.xlsx')
      expect(noteWorkbookIdentity('/books/budget.xlsx')).toBe(false)
    })

    it('reports a switch and drops the previous workbook drift', () => {
      open('/books/a.xlsx')
      noteWorkbookMutation(INSERT_ROW, false)
      expect(workbookDrift().structural).toBe(true)
      expect(noteWorkbookIdentity('/books/b.xlsx')).toBe(true)
      // Those counters described the file that was just closed; the new one has
      // no edits of its own.
      expect(workbookDrift()).toEqual({ structural: false, content: false })
      expect(noteWorkbookIdentity('/books/b.xlsx')).toBe(false)
    })

    it('treats an unknown identity as "no information", not as a switch', () => {
      open('/books/known.xlsx')
      // A renderer that cannot name the workbook must not invent a switch — the
      // note tells the model its whole conversation is invalid…
      expect(noteWorkbookIdentity(null)).toBe(false)
      // …and the last known identity survives, so a real switch is still caught.
      expect(noteWorkbookIdentity('/books/other.xlsx')).toBe(true)
    })
  })

  describe('noteWorkbookDocumentReplaced: the loader says a new document took over', () => {
    it('reports once, on the turn after the load', () => {
      noteWorkbookIdentity('/books/live.xlsx')
      noteWorkbookDocumentReplaced()
      expect(noteWorkbookIdentity('/books/live.xlsx')).toBe(true)
      // Consumed: the switch is history by the second turn, not a standing warning.
      expect(noteWorkbookIdentity('/books/live.xlsx')).toBe(false)
    })

    it('reports even when the new document has no path to compare', () => {
      // The case the path comparison structurally cannot see: a new unsaved
      // workbook (identity null) replaced by another one.
      noteWorkbookDocumentReplaced()
      expect(noteWorkbookIdentity(null)).toBe(true)
      expect(noteWorkbookIdentity(null)).toBe(false)
    })

    it('drops the previous document drift along with it', () => {
      noteWorkbookIdentity('/books/one.xlsx')
      markWorkbookSeen()
      noteWorkbookMutation(INSERT_ROW, false)
      expect(workbookDrift().structural).toBe(true)
      noteWorkbookDocumentReplaced()
      expect(noteWorkbookIdentity('/books/two.xlsx')).toBe(true)
      // Those counts described the document that was just closed; no markWorkbookSeen
      // needed — the baseline went with it.
      expect(workbookDrift()).toEqual({ structural: false, content: false })
    })
  })

  describe('realignWorkbookIdentity: the same document under a new path', () => {
    const open = (path: string): void => {
      noteWorkbookIdentity(path)
      markWorkbookSeen()
    }

    it('excuses a Save As, and keeps the edits that describe the document', () => {
      open('/books/draft.xlsx')
      noteWorkbookMutation(SET_VALUES, false)
      // Save As reopens through the same "session swap over the same document"
      // path a plain save uses, so the path moves while the document does not.
      realignWorkbookIdentity('/books/final.xlsx')
      expect(noteWorkbookIdentity('/books/final.xlsx')).toBe(false)
      expect(workbookDrift().content).toBe(true)
    })

    it('still reports a switch that follows one', () => {
      open('/books/one.xlsx')
      realignWorkbookIdentity('/books/two.xlsx')
      expect(noteWorkbookIdentity('/books/three.xlsx')).toBe(true)
    })

    it('does not invent an identity when the path is unknown', () => {
      open('/books/keep.xlsx')
      realignWorkbookIdentity(null)
      // Losing the path (a converted import, a closed file) must not erase the
      // last known identity, or the next real switch goes unreported.
      expect(noteWorkbookIdentity('/books/keep.xlsx')).toBe(false)
      expect(noteWorkbookIdentity('/books/gone.xlsx')).toBe(true)
    })
  })

  describe('first conversation', () => {
    it('nothing is stale before the model has been shown anything', () => {
      // buildContext marks the baseline before the first turn, so a fresh panel
      // must never open by refusing a write.
      expect(workbookDrift()).toEqual({ structural: false, content: false })
    })
  })
})
