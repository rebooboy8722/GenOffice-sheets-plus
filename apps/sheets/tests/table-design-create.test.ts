import { describe, expect, it, vi } from 'vitest'

import { createEditJournal } from '../src/renderer/edit-journal'
import { t } from '../src/renderer/i18n/locale'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'
// Static import (not `await import` inside the test case): the module graph's
// ~11s cold transform would otherwise land inside the 5s per-test timeout.
// The vi.mock calls above are hoisted before it, so the mocks still apply.
import { handleTableDesignCommand } from '../src/renderer/table-design-actions'

vi.mock('../src/renderer/univer-sync', () => ({
  ensureLazyRangeLoaded: vi.fn(() => Promise.resolve(true)),
  loadVisibleRange: vi.fn(() => Promise.resolve()),
}))
vi.mock('../src/renderer/workbook-ops', () => ({
  applyAiTableAdd: vi.fn(),
  applySessionTableTheme: vi.fn(),
}))

const flush = () => new Promise((resolve) => setTimeout(resolve, 0))

function fakeSheet(rows: number) {
  const activated: number[][] = []
  const range = {
    getValues: () => [[null, null]],
    setValue: () => range,
    setValues: () => range,
    clearContent: () => range,
    activate: () => activated.push([1]),
  }
  const worksheet = {
    getSheetId: () => 's1',
    getRange: () => range,
    getLastRow: () => rows - 1,
    getLastColumn: () => 1,
  }
  const runtime = {
    univerAPI: {
      getActiveWorkbook: () => ({
        getActiveSheet: () => worksheet,
        getActiveRange: () => null,
      }),
    },
  } as unknown as UniverRuntime
  const state = {
    file: { sheets: [{ id: 's1', name: 'Data', tables: [], rowCount: rows, columnCount: 2 }] },
    editJournal: createEditJournal(),
    loadedRanges: new Map(),
    frozenStripKeys: new Set(),
  } as unknown as LazyWorkbookState
  return { runtime, state, activated }
}

describe('Create Table completion signal', () => {
  it('fires onTableCreated after the asynchronous create succeeds', async () => {
    const { runtime, state, activated } = fakeSheet(10)
    const created = vi.fn()
    const messages: string[] = []
    handleTableDesignCommand(
      {
        univerRef: { current: runtime },
        lazyWorkbookRef: { current: state },
        setMessage: (message) => messages.push(message),
        setPendingEdits: () => undefined,
        refreshSelection: () => undefined,
        onTableCreated: created,
      },
      'table-create:A1:B4:1',
    )
    expect(created).not.toHaveBeenCalled()
    await flush()
    expect(created).toHaveBeenCalledTimes(1)
    expect(activated).toHaveLength(1)
    expect(messages).toEqual([t('appTableCreated')])
  })

  it('does not signal when the create fails', async () => {
    const { runtime, state } = fakeSheet(10)
    const created = vi.fn()
    handleTableDesignCommand(
      {
        univerRef: { current: runtime },
        lazyWorkbookRef: { current: state },
        setMessage: () => undefined,
        setPendingEdits: () => undefined,
        refreshSelection: () => undefined,
        onTableCreated: created,
      },
      'table-create:not-a-range:1',
    )
    await flush()
    expect(created).not.toHaveBeenCalled()
  })
})
