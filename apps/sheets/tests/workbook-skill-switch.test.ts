import { describe, expect, it } from 'vitest'

import { createWorkbookSkill } from '../src/renderer/ai/workbook-skill'
import { WORKBOOK_SWITCHED_NOTE } from '../src/renderer/ai/workbook-baseline'
import type { SheetsSkillDeps } from '../src/renderer/ai/tools'

/**
 * The one signal that a conversation has been left pointing at another file.
 *
 * Opening a different workbook reuses this renderer: the chat is not cleared
 * (`AgentLoop.reset()` runs only on "New chat") while every range the model read
 * earlier belonged to the closed file. No drift counter can see that — nothing
 * was edited, the workbook was replaced — so `buildContext` has to say so.
 *
 * The overview itself is not the subject here (mode `none` keeps the fake deps
 * honest); the subject is which header the skill puts in front of it.
 */
function skillWithIdentity(getId: () => string | null) {
  const deps = {
    getActiveSheetInfo: () => ({ mode: 'none' as const }),
    workbookIdentity: getId,
  } as unknown as SheetsSkillDeps
  return createWorkbookSkill(deps).buildContext!
}

describe('workbook skill: a replaced workbook is announced', () => {
  it('says nothing on the first turn', () => {
    let id: string | null = 'book-1'
    const build = skillWithIdentity(() => id)
    const first = build()
    expect(first).toContain('No workbook is currently open')
    expect(first).not.toContain(WORKBOOK_SWITCHED_NOTE.slice(0, 40))
    id = 'book-1'
    expect(build()).not.toContain('DIFFERENT WORKBOOK')
  })

  it('warns on the turn after another workbook is opened', () => {
    let id: string | null = 'book-1'
    const build = skillWithIdentity(() => id)
    build()
    id = 'book-2'
    const text = build()
    expect(text).toContain('DIFFERENT WORKBOOK IS NOW OPEN')
    expect(text).toContain('Do NOT reuse those addresses')
    // The warning leads: a model that skims the first line still gets it.
    expect(text.indexOf('DIFFERENT WORKBOOK')).toBeLessThan(
      text.indexOf('No workbook is currently open'),
    )
  })

  it('warns once, not on every later turn', () => {
    let id: string | null = 'book-a'
    const build = skillWithIdentity(() => id)
    build()
    id = 'book-b'
    expect(build()).toContain('DIFFERENT WORKBOOK')
    expect(build()).not.toContain('DIFFERENT WORKBOOK')
  })

  it('treats an unnamed workbook as "no information", not as a switch', () => {
    // Closing the last file leaves the renderer with no workbook to name. That
    // must not claim a switch — the note tells the model its whole conversation
    // is invalid — and the last known identity has to survive the gap, so
    // re-opening the SAME file is still not a switch. (A real switch after a gap
    // is covered in tests/workbook-baseline.test.ts.)
    let id: string | null = 'book-keep'
    const build = skillWithIdentity(() => id)
    build()
    id = null
    expect(build()).not.toContain('DIFFERENT WORKBOOK')
    expect(build()).not.toContain('DIFFERENT WORKBOOK')
    id = 'book-keep'
    expect(build()).not.toContain('DIFFERENT WORKBOOK')
  })
})
