/**
 * The throttle's contract, and the one bug it must never ship: a frame
 * scheduled while the model was streaming must not land after the run ended,
 * or half-finished prose overwrites the completion (or error) text the terminal
 * event wrote.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'

import { createStreamCoalescer, type StreamPending } from '../src/renderer/ai/stream-coalescer'

/** Stand-in for rAF: frames are fired by the test, so ordering is explicit. */
function fakeScheduler(): {
  request(cb: () => void): number
  cancel(handle: number): void
  /** runs every frame still queued, as the browser would at the next paint */
  paint(): void
  queued(): number
} {
  let next = 1
  const frames = new Map<number, () => void>()
  return {
    request: (cb) => {
      const id = next++
      frames.set(id, cb)
      return id
    },
    cancel: (handle) => {
      frames.delete(handle)
    },
    paint: () => {
      const due = [...frames.entries()]
      frames.clear()
      for (const [, cb] of due) cb()
    },
    queued: () => frames.size,
  }
}

describe('stream coalescer', () => {
  let commits: StreamPending[]
  let frames: ReturnType<typeof fakeScheduler>
  let stream: ReturnType<typeof createStreamCoalescer>

  beforeEach(() => {
    commits = []
    frames = fakeScheduler()
    stream = createStreamCoalescer((pending) => commits.push(pending), frames)
  })

  it('collapses a burst into one commit carrying the newest text', () => {
    stream.push('text', 'a')
    stream.push('text', 'ab')
    stream.push('text', 'abc')
    expect(commits).toHaveLength(0)
    expect(frames.queued()).toBe(1)

    frames.paint()
    expect(commits).toEqual([{ text: 'abc' }])
  })

  it('keeps thinking and prose on independent clocks within one frame', () => {
    // A reasoning model streams thinking for seconds before any prose; neither
    // may hold the other back, and both must survive the same coalesced commit.
    stream.push('reasoning', 'think')
    stream.push('text', 'hello')
    frames.paint()
    expect(commits).toEqual([{ text: 'hello', reasoning: 'think' }])
  })

  it('starts a new frame after one lands', () => {
    stream.push('text', 'one')
    frames.paint()
    stream.push('text', 'two')
    frames.paint()
    expect(commits).toEqual([{ text: 'one' }, { text: 'two' }])
  })

  it('commits immediately on flush and schedules nothing further', () => {
    stream.push('text', 'narration')
    stream.flush()
    expect(commits).toEqual([{ text: 'narration' }])
    expect(frames.queued()).toBe(0)
    frames.paint()
    expect(commits).toHaveLength(1)
  })

  it('does not commit an empty frame', () => {
    stream.flush()
    stream.discard()
    frames.paint()
    expect(commits).toHaveLength(0)
  })

  it('a discarded frame never commits, even when it still fires', () => {
    stream.push('text', 'half-finished prose')
    stream.discard()
    // The release cleared the queue, so a paint is a no-op…
    frames.paint()
    expect(commits).toHaveLength(0)
  })

  it('a frame that fires late cannot resurrect discarded text', () => {
    // The hostile version: the scheduler ignores the cancel (a browser never
    // does, but the coalescer must not depend on that to stay correct).
    const deaf = fakeScheduler()
    const commitsLate: StreamPending[] = []
    const guarded = createStreamCoalescer((p) => commitsLate.push(p), {
      request: deaf.request,
      cancel: () => {},
    })
    guarded.push('text', 'stale')
    guarded.discard()
    deaf.paint()
    expect(commitsLate).toHaveLength(0)
  })

  it('resumes normally after a discard', () => {
    stream.push('text', 'stale')
    stream.discard()
    stream.push('text', 'final-run')
    frames.paint()
    expect(commits).toEqual([{ text: 'final-run' }])
  })

  it('flushes only what arrived, so a later turn is not committed twice', () => {
    stream.push('text', 'turn one')
    stream.flush()
    stream.flush()
    expect(commits).toEqual([{ text: 'turn one' }])
  })

  it('clears reasoning once shown, so a text-only frame does not re-commit it', () => {
    stream.push('reasoning', 'thinking')
    frames.paint()
    stream.push('text', 'answer')
    frames.paint()
    expect(commits).toEqual([{ reasoning: 'thinking' }, { text: 'answer' }])
  })

  it('leaves the commit decision to the host: no push, no commit call', () => {
    const spy = vi.fn()
    const idle = createStreamCoalescer(spy, frames)
    idle.flush()
    idle.discard()
    frames.paint()
    expect(spy).not.toHaveBeenCalled()
  })
})
