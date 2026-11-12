// @vitest-environment jsdom
/**
 * The thinking block's open/closed behaviour is expressed by one prop spread:
 * `open` is passed only while there is no answer yet. That relies on React
 * actually closing the disclosure when the attribute goes away — a claim about
 * React's DOM handling, so it is tested rather than asserted in a comment.
 *
 * Rendered as an isolated stand-in for the panel's block: the behaviour under
 * test lives entirely in this attribute pattern, and pulling in AiChatPanel
 * would need the whole i18n/desktopApi/Univer surface around it.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

/** Mirrors apps/sheets/src/renderer/ai/AiChatPanel.tsx's reasoning block. */
function ReasoningBlock(props: { reasoning?: string; text?: string }) {
  const { reasoning, text } = props
  if (!reasoning) return null
  return createElement(
    'details',
    { className: 'ai-reasoning', ...(text ? {} : { open: true }) },
    createElement('summary', null, text ? 'Reasoning' : 'Thinking…'),
    createElement('div', { className: 'ai-reasoning-body', dir: 'auto' }, reasoning),
  )
}

let root: Root | null = null
let container: HTMLDivElement | null = null

const details = (): HTMLDetailsElement | null =>
  container?.querySelector('details.ai-reasoning') ?? null

async function render(props: { reasoning?: string; text?: string }): Promise<void> {
  await act(async () => {
    root?.render(createElement(ReasoningBlock, props))
  })
}

beforeEach(() => {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
})

describe('reasoning disclosure', () => {
  it('is open while thinking is all there is, so the wait does not read as a hang', async () => {
    await render({ reasoning: 'let me check the column…' })
    expect(details()?.open).toBe(true)
    expect(container?.querySelector('summary')?.textContent).toBe('Thinking…')
    expect(container?.textContent).toContain('let me check the column')
  })

  it('grows with each delta in place', async () => {
    await render({ reasoning: 'think' })
    await render({ reasoning: 'thinking further' })
    expect(details()?.open).toBe(true)
    expect(container?.querySelector('.ai-reasoning-body')?.textContent).toBe('thinking further')
  })

  it('collapses once the answer arrives', async () => {
    await render({ reasoning: 'done thinking' })
    expect(details()?.open).toBe(true)

    await render({ reasoning: 'done thinking', text: 'Here is the result.' })
    // The whole point: the answer must not be pushed below an open wall of thinking.
    // (The answer itself is rendered by the panel beside this block, so only the
    // disclosure's own state is in scope here — and the label switches tense.)
    expect(details()?.open).toBe(false)
    expect(container?.querySelector('summary')?.textContent).toBe('Reasoning')
  })

  it('stays collapsed as the answer streams on', async () => {
    await render({ reasoning: 'thoughts', text: 'part one' })
    await render({ reasoning: 'thoughts', text: 'part one, part two' })
    expect(details()?.open).toBe(false)
  })

  it('a collapsed block keeps its text for re-reading', async () => {
    await render({ reasoning: 'the reasoning to revisit', text: 'answer' })
    const body = container?.querySelector('.ai-reasoning-body')
    expect(body?.textContent).toBe('the reasoning to revisit')
    // Hidden by the disclosure, not removed from the document.
    expect(body).toBeTruthy()
  })

  it('disappears entirely when the turn had no thinking', async () => {
    await render({ text: 'a plain non-reasoning answer' })
    expect(details()).toBeNull()
  })

  it('drops the block when the turn ends and thinking is cleared', async () => {
    await render({ reasoning: 'turn-local thoughts', text: 'answer' })
    expect(details()).toBeTruthy()
    // onTurnEnd clears reasoning; the record must not linger on a finished message.
    await render({ text: 'answer' })
    expect(details()).toBeNull()
  })
})
