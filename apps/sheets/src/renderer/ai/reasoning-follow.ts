import { useLayoutEffect, useRef, type RefObject } from 'react'

/**
 * Tail-follow for the live thinking box.
 *
 * The thinking text lives in a height-capped scroll box (`.ai-reasoning-body`,
 * `max-height: 220px` in styles.css — the cap keeps a long chain of thought
 * from pushing the reply out of the panel). That cap is exactly why the newest
 * thinking needs its own follower: growth happens *inside* the box, so the
 * transcript's stick-to-bottom never fires for it, and without help the view
 * sits at whatever was read first while the model's latest lines pile up below
 * the fold — the scrollbar thumb shrinking is the only hint anything is moving.
 *
 * The rule mirrors the transcript's `onChatScroll`: follow while the reader is
 * at (or near) the bottom, stop yanking once they scroll up to re-read, resume
 * when they come back down. Follow from the start: the box begins "at the
 * bottom" (nothing has overflowed yet), and a reader who never touches the
 * wheel sees every line as it arrives.
 */

/** Within this distance of the bottom the reader still counts as "at the latest line" — same rule and value as the transcript's follow. */
const FOLLOW_THRESHOLD_PX = 48

export interface ReasoningFollow {
  /** Attach to the scroll box element. */
  readonly ref: RefObject<HTMLDivElement | null>
  /** Attach to the same element's `onScroll`. */
  readonly onScroll: () => void
}

export function useReasoningFollow(text: string): ReasoningFollow {
  const ref = useRef<HTMLDivElement | null>(null)
  const stickRef = useRef(true)
  // layout, not passive: scrolling before paint means the reader never sees a
  // frame of the old position between two deltas
  useLayoutEffect(() => {
    const el = ref.current
    if (el && stickRef.current) el.scrollTop = el.scrollHeight
  }, [text])
  return {
    ref,
    onScroll: () => {
      const el = ref.current
      if (el) {
        stickRef.current = el.scrollHeight - el.scrollTop - el.clientHeight < FOLLOW_THRESHOLD_PX
      }
    },
  }
}
