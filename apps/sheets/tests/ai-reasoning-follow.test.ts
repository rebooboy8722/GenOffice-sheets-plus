// @vitest-environment jsdom
/**
 * The thinking box must chase its own bottom — the user's complaint was that
 * the thinking stream grows while the view stays wherever it was, the newest
 * lines piling up below the 220 px fold (and the scrollbar thumb shrinking is
 * the only hint anything is moving at all).
 *
 * These tests import the **real** `useReasoningFollow` (not a mirror — a
 * mirror stays green while the source rots, see ai-reasoning-disclosure's own
 * comment for why it is limited to the attribute pattern). jsdom has no
 * layout, so the scroll box's geometry is faked per element: scrollHeight /
 * clientHeight as readable properties and scrollTop as a recorded assignment,
 * which is exactly what the hook reads and writes.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { useReasoningFollow } from '../src/renderer/ai/reasoning-follow'

function Host({ text }: { text: string }) {
  const { ref, onScroll } = useReasoningFollow(text)
  return createElement('div', { ref, onScroll }, text)
}

interface FakeBox {
  scrollTop: number
  setScrollHeight(v: number): void
}

/** Fake the layout the hook observes; assignments to scrollTop are recorded. */
function fakeBox(el: HTMLElement, scrollHeight: number, clientHeight: number): FakeBox {
  const state = { scrollTop: 0, scrollHeight, clientHeight }
  Object.defineProperty(el, 'scrollHeight', { configurable: true, get: () => state.scrollHeight })
  Object.defineProperty(el, 'clientHeight', { configurable: true, get: () => state.clientHeight })
  Object.defineProperty(el, 'scrollTop', {
    configurable: true,
    get: () => state.scrollTop,
    set: (v: number) => {
      state.scrollTop = v
    },
  })
  return {
    get scrollTop() {
      return state.scrollTop
    },
    set scrollTop(v: number) {
      state.scrollTop = v
    },
    setScrollHeight(v: number) {
      state.scrollHeight = v
    },
  }
}

/** The reader's wheel: move the box and let React see the scroll event. */
function scrollByReader(el: HTMLElement, box: FakeBox, top: number): void {
  box.scrollTop = top
  el.dispatchEvent(new Event('scroll'))
}

let root: Root | null = null
let container: HTMLDivElement | null = null

async function render(text: string): Promise<void> {
  await act(async () => {
    root?.render(createElement(Host, { text }))
  })
}

function host(): HTMLElement {
  const el = container?.querySelector('div')
  if (!el) throw new Error('host not rendered')
  return el
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

describe('reasoning tail-follow', () => {
  it('follows the bottom from the start — the reader never touched the wheel', async () => {
    await render('think')
    const box = fakeBox(host(), 300, 220)
    // the box overflowed (300 > 220) without the user doing anything: the
    // next delta must land at the bottom, not at the top where it started
    await render('thinking further')
    expect(box.scrollTop).toBe(300)
  })

  it('keeps following every delta while nobody scrolls away', async () => {
    await render('a')
    const box = fakeBox(host(), 300, 220)
    await render('ab')
    expect(box.scrollTop).toBe(300)
    box.setScrollHeight(500)
    await render('abc')
    expect(box.scrollTop).toBe(500)
  })

  it('stops yanking once the reader scrolls up to re-read', async () => {
    await render('a')
    const el = host()
    const box = fakeBox(el, 300, 220)
    await render('ab')
    expect(box.scrollTop).toBe(300)
    // 70 px from the bottom: reading history, not following
    scrollByReader(el, box, 10)
    box.setScrollHeight(500)
    await render('abc')
    expect(box.scrollTop).toBe(10) // the history stays put
  })

  it('resumes when the reader comes back to the bottom', async () => {
    await render('a')
    const el = host()
    const box = fakeBox(el, 300, 220)
    scrollByReader(el, box, 10)
    box.setScrollHeight(500)
    await render('ab')
    expect(box.scrollTop).toBe(10)
    // back within the threshold (500 − 280 − 220 = 0)
    scrollByReader(el, box, 280)
    box.setScrollHeight(700)
    await render('abc')
    expect(box.scrollTop).toBe(700)
  })

  it('the threshold is the same 48 px rule as the transcript', async () => {
    await render('a')
    const el = host()
    const box = fakeBox(el, 700, 220)
    // 40 px from the bottom → still "at the latest line"
    scrollByReader(el, box, 440)
    box.setScrollHeight(900)
    await render('ab')
    expect(box.scrollTop).toBe(900)
    // 60 px → scrolled away
    scrollByReader(el, box, 620)
    box.setScrollHeight(1100)
    await render('abc')
    expect(box.scrollTop).toBe(620)
  })
})
