import type { AgentSkill } from '@genoffice/agent-core'
import basePrompt from './prompts/base.md?raw'
import { createDestructiveGateState } from './destructive-guard'
import { verifySheetsResponse } from './response-verify'
import {
  WORKBOOK_TOOLS,
  buildWorkbookContext,
  executeWorkbookTool,
  type SheetsSkillDeps,
} from './tools'
import {
  gateWorkbookWrite,
  markWorkbookSeen,
  noteWorkbookIdentity,
  WORKBOOK_SWITCHED_NOTE,
} from './workbook-baseline'

/**
 * The workbook DSL as an AgentSkill: mirrors createDocsSkill's shape
 * (systemPrompt + tools + buildContext + executeTool) so it plugs into the
 * same packages/agent-core AgentLoop docx uses.
 *
 * Prompt layout: the always-loaded base prompt (prompts/base.md) stays small
 * — workflow, op catalog, cross-cutting discipline — while per-domain field
 * definitions and conventions live in prompts/guides/*.md, loaded on demand
 * via load_guide.
 */
export function createWorkbookSkill(deps: SheetsSkillDeps): AgentSkill {
  // F16: the destructive-confirmation gate is session-scoped (one per skill
  // instance). The loop ticks it once per user run (onRunStart); a flagged
  // resubmission is only honored after a refusal in an EARLIER run, so a
  // preemptive flag or a same-turn retry can no longer bypass the round trip.
  const destructiveGate = createDestructiveGateState()
  const gatedDeps: SheetsSkillDeps = { ...deps, destructiveGate }
  return {
    id: 'sheets',
    onRunStart: () => {
      destructiveGate.runSeq += 1
    },
    systemPrompt: basePrompt,
    tools: WORKBOOK_TOOLS,
    // The context this returns is the freshness baseline for range-addressed
    // writes: whatever the model is shown here is what it will address against.
    buildContext: () => {
      // Identity first: when the file was replaced under this conversation, the
      // previous workbook's edit counts are meaningless (and would read as drift
      // of the new one), and the model has to be told — no counter can see a
      // workbook being swapped for another.
      const switched = noteWorkbookIdentity(deps.workbookIdentity?.() ?? null)
      const context = buildWorkbookContext(deps)
      markWorkbookSeen()
      return switched ? `${WORKBOOK_SWITCHED_NOTE}\n\n${context}` : context
    },
    executeTool: (call) =>
      gateWorkbookWrite(
        call.name,
        (summary, output) => ({ output, isError: true, mutated: false, summary }),
        () => executeWorkbookTool(call, gatedDeps),
      ),
    verifyResponse: verifySheetsResponse,
  }
}
