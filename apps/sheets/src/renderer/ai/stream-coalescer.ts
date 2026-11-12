/**
 * Coalesces streamed model output into at most one commit per animation frame.
 *
 * `onText` / `onReasoning` hand over the turn's *accumulated* text on every SSE
 * delta, and each commit re-renders the whole app tree (App → ExcelShell →
 * panel) plus a full Markdown re-parse of everything said so far. At a local
 * endpoint's ~130 tok/s that is 130 full renders per second fighting Univer's
 * grid for the main thread — the faster the endpoint, the worse it reads.
 *
 * rAF rather than a fixed timer: it self-adapts (a frame over budget simply
 * commits less often), and the newest value always wins, so nothing is lost —
 * only delayed by under a frame.
 *
 * The invariant worth guarding: **a frame scheduled before a run ended must
 * never land after it**. The run's terminal events (`onDone` / `onError`) write
 * the final text themselves, so a stale pending delta would overwrite the
 * completion or error message with half-finished prose. Hence `discard` on the
 * terminal path and `flush` wherever another event writes the message first.
 */

/** Which streamed field a value belongs to. Independent so a reasoning model's
 * thinking (which streams for seconds before any prose) never blocks its text. */
export type StreamField = 'text' | 'reasoning'

export type StreamPending = Partial<Record<StreamField, string>>

interface Scheduler {
  request(cb: () => void): number
  cancel(handle: number): void
}

const frameScheduler: Scheduler = {
  request: (cb) => requestAnimationFrame(cb),
  cancel: (handle) => cancelAnimationFrame(handle),
}

export interface StreamCoalescer {
  /** Records the latest value of one field; schedules the frame if idle. */
  push(field: StreamField, value: string): void
  /** Commits now (an unrelated event is about to write the message). */
  flush(): void
  /** Drops what is pending — the caller is writing the run's final text. */
  discard(): void
}

export function createStreamCoalescer(
  commit: (pending: StreamPending) => void,
  scheduler: Scheduler = frameScheduler,
): StreamCoalescer {
  let pending: StreamPending = {}
  let frame: number | null = null

  const settle = () => {
    if (frame === null) return
    scheduler.cancel(frame)
    frame = null
  }
  const drain = () => {
    settle()
    if (pending.text === undefined && pending.reasoning === undefined) return
    const ready = pending
    pending = {}
    commit(ready)
  }

  return {
    push(field, value) {
      pending[field] = value
      if (frame !== null) return
      frame = scheduler.request(() => {
        frame = null
        drain()
      })
    },
    flush: drain,
    discard() {
      settle()
      pending = {}
    },
  }
}
