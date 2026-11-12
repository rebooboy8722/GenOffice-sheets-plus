/**
 * An explicit hold covers one bulk write end to end. The time-window hold is
 * driven by chunk arrivals tens of milliseconds apart, so it cannot serve a
 * writer that awaits between ops (an AI apply prefetches images and pins
 * streamed precedents): its ops are seconds apart, the window expires between
 * them, and the engine rebuilds the whole dependency tree once per op. Here a
 * held batch must produce zero cycles, and its release exactly one.
 */
import {
  ICommandService,
  IUniverInstanceService,
  LocaleType,
  LogLevel,
  Univer,
  UniverInstanceType,
} from '@univerjs/core'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import { describe, expect, it } from 'vitest'

import {
  FORMULA_STREAM_HOLD_MS,
  holdFormulaCalculation,
  installFormulaStreamHold,
} from '../src/renderer/formula-stream-hold'

const START_MUTATION = 'formula.mutation.set-formula-calculation-start'

function setupWorkbook(): {
  univer: Univer
  runtime: Record<string, unknown>
  commandService: ICommandService
  /** calculation cycles that actually reached the engine */
  started: number[]
  setRow(row: number, value: number): Promise<unknown>
  sum(): number | undefined
  settle(): Promise<void>
} {
  const univer = new Univer({ logLevel: LogLevel.ERROR, locale: LocaleType.EN_US, locales: {} })
  univer.registerPlugin(UniverFormulaEnginePlugin)
  univer.registerPlugin(UniverSheetsPlugin)
  univer.registerPlugin(UniverSheetsFormulaPlugin)
  const injector = univer.__getInjector()
  const runtime = { univer, univerAPI: null as never }
  installFormulaStreamHold(runtime as never)
  univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
    id: 'wb1',
    sheetOrder: ['s1'],
    name: 'wb',
    styles: {},
    sheets: { s1: { id: 's1', name: 'S1', rowCount: 10, columnCount: 5, cellData: {} } },
  })
  const commandService = injector.get(ICommandService)
  const started: number[] = []
  commandService.onCommandExecuted((command) => {
    if (command.id === START_MUTATION) started.push(started.length)
  })
  const workbook = injector.get(IUniverInstanceService).getUnit('wb1') as unknown as {
    getSheetBySheetId(id: string): { getCellRaw(row: number, column: number): { v?: unknown } }
  }
  return {
    univer,
    runtime,
    commandService,
    started,
    // One cell per command: each mutation is what would otherwise kick off its
    // own cycle, which is exactly the cost being removed.
    setRow: (row, value) =>
      commandService.executeCommand('sheet.command.set-range-values', {
        unitId: 'wb1',
        subUnitId: 's1',
        range: { startRow: row, endRow: row, startColumn: 0, endColumn: 0 },
        value: { [row]: { 0: { v: value } } },
      }),
    sum: () => workbook.getSheetBySheetId('s1').getCellRaw(0, 1)?.v as number | undefined,
    settle: () => new Promise((resolve) => setTimeout(resolve, FORMULA_STREAM_HOLD_MS + 600)),
  }
}

/** Installs `B1 = SUM(A2:A4)` and lets its first cycle land, so the counter and
 * the dependent both start from a computed baseline (an uncalculated formula
 * reads as `undefined`, not 0, and would make the held-state assertion vacuous). */
async function baseline(
  univer: Univer,
  started: number[],
  settle: () => Promise<void>,
): Promise<void> {
  await installSumFormula(univer)
  await settle()
  expect(started.length).toBeGreaterThan(0)
  started.length = 0
}

describe('holdFormulaCalculation', () => {
  it('is a no-op for a runtime the hold was never installed on', () => {
    const release = holdFormulaCalculation({ univer: null } as never)
    expect(() => release()).not.toThrow()
    expect(() => release()).not.toThrow()
  })

  it('batches a held write into one cycle and still computes the final value', async () => {
    const { univer, runtime, started, setRow, sum, settle } = setupWorkbook()
    await baseline(univer, started, settle)

    const release = holdFormulaCalculation(runtime as never)
    for (const [row, value] of [
      [1, 2],
      [2, 3],
      [3, 4],
    ] as const) {
      await setRow(row, value)
    }
    // Held: the engine never got a kick-off, so the sum is still the stale one.
    expect(started).toHaveLength(0)
    expect(sum()).toBe(0)

    release()
    await settle()
    expect(started).toHaveLength(1)
    expect(sum()).toBe(9)
    univer.dispose()
  }, 20_000)

  it('releases only when the last nested holder lets go', async () => {
    const { univer, runtime, started, setRow, sum, settle } = setupWorkbook()
    await baseline(univer, started, settle)

    const releaseOuter = holdFormulaCalculation(runtime as never)
    const releaseInner = holdFormulaCalculation(runtime as never)
    await setRow(1, 5)
    releaseInner()
    await settle()
    // The outer hold is still open: nothing may reach the engine yet.
    expect(started).toHaveLength(0)
    expect(sum()).toBe(0)

    await setRow(2, 5)
    releaseOuter()
    await settle()
    expect(started).toHaveLength(1)
    expect(sum()).toBe(10)
    univer.dispose()
  }, 20_000)
})

/** B1 = SUM(A2:A4) — a dependent of every cell the tests write below it. */
async function installSumFormula(univer: Univer): Promise<void> {
  const commandService = univer.__getInjector().get(ICommandService)
  await commandService.executeCommand('sheet.command.set-range-values', {
    unitId: 'wb1',
    subUnitId: 's1',
    range: { startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 },
    value: { 0: { 1: { f: '=SUM(A2:A4)' } } },
  })
}
