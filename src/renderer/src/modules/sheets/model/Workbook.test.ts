import { describe, expect, it } from 'vitest'
import { blankWorkbookData, Workbook } from './Workbook'
import type { WorkbookData } from './workbook-data'

/** Four regions with revenue, cost, and a per-row margin formula; totals on row 6. */
function regionsData(): WorkbookData {
  const blank = blankWorkbookData()
  const sheet = blank.sheets[0]
  if (sheet === undefined) throw new Error('Blank workbook has no sheet.')
  return {
    ...blank,
    sheets: [{
      ...sheet,
      cells: [
        ['Region', 'Revenue', 'Cost', 'Margin'],
        ['North', 42100, 18300, '=IFERROR((B2-C2)/B2,0)'],
        ['South', 38750, 21900, '=IFERROR((B3-C3)/B3,0)'],
        ['East', 29400, 14050, '=IFERROR((B4-C4)/B4,0)'],
        ['West', 33860, 19240, '=IFERROR((B5-C5)/B5,0)'],
        ['Total', '=SUM(B2:B5)', '=SUM(C2:C5)', '=IFERROR((B6-C6)/B6,0)'],
      ],
    }],
  }
}

function regionsWorkbook(): Workbook {
  return new Workbook(regionsData())
}

describe('Workbook', () => {
  it('recalculates dependent cells after an edit', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id
    if (sheetId === undefined) throw new Error('Starter sheet is missing.')

    workbook.setCellInput(sheetId, [1, 1], 50_000)

    expect(workbook.getCellValue(sheetId, [1, 5])).toBe(152_010)
  })

  it('calculates references across added and renamed sheets', () => {
    const workbook = regionsWorkbook()
    const firstSheet = workbook.sheets()[0]
    if (firstSheet === undefined) throw new Error('Starter sheet is missing.')
    const secondSheet = workbook.addSheet()
    workbook.renameSheet(secondSheet.id, 'Assumptions')

    workbook.setCellInput(secondSheet.id, [0, 0], 1.25)
    workbook.setCellInput(firstSheet.id, [5, 0], '=Assumptions!A1*100')

    expect(workbook.getCellValue(firstSheet.id, [5, 0])).toBe(125)
  })

  it('shifts formulas during structural edits and restores them through undo', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id
    if (sheetId === undefined) throw new Error('Starter sheet is missing.')

    workbook.insertRows(sheetId, 1)

    expect(workbook.getCellInput(sheetId, [1, 6])).toBe('=SUM(B3:B6)')
    expect(workbook.getCellValue(sheetId, [1, 6])).toBe(144_110)

    workbook.undo()
    const restoredSheetId = workbook.sheets()[0]?.id
    if (restoredSheetId === undefined) throw new Error('Restored sheet is missing.')
    expect(workbook.getCellInput(restoredSheetId, [1, 5])).toBe('=SUM(B2:B5)')
  })

  it('infers values and number formats from spreadsheet clipboard text', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id
    if (sheetId === undefined) throw new Error('Starter sheet is missing.')

    workbook.setBlock(sheetId, [5, 1], [['$1,240.50', '37.5%', '2026-08-13']])

    expect(workbook.getCellValue(sheetId, [5, 1])).toBe(1_240.5)
    expect(workbook.getStyle(sheetId, [5, 1]).numberFormat).toBe('currency')
    expect(workbook.getStyle(sheetId, [5, 1]).decimalPlaces).toBe(2)
    expect(workbook.getCellValue(sheetId, [6, 1])).toBe(0.375)
    expect(workbook.getStyle(sheetId, [6, 1]).numberFormat).toBe('percent')
    expect(workbook.getStyle(sheetId, [7, 1]).numberFormat).toBe('date')
  })

  it('keeps format, structure, and sheet mutations in one undo history', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id
    if (sheetId === undefined) throw new Error('Starter sheet is missing.')

    workbook.setNumberFormat(sheetId, { x: 0, y: 0, width: 1, height: 1 }, 'currency')
    workbook.insertColumns(sheetId, 0)
    workbook.addSheet()

    workbook.undo()
    expect(workbook.sheets()).toHaveLength(1)
    workbook.undo()
    const restoredSheetId = workbook.sheets()[0]?.id
    if (restoredSheetId === undefined) throw new Error('Restored sheet is missing.')
    expect(workbook.getColumnCount(restoredSheetId)).toBe(26)
    workbook.undo()
    expect(workbook.getStyle(restoredSheetId, [0, 0]).numberFormat).toBe('general')

    workbook.redo()
    expect(workbook.getStyle(restoredSheetId, [0, 0]).numberFormat).toBe('currency')
  })

  it('translates formulas and propagates formats when filling a range', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id
    if (sheetId === undefined) throw new Error('Starter sheet is missing.')
    workbook.setCellInput(sheetId, [0, 9], 2)
    workbook.setCellInput(sheetId, [0, 10], 3)
    workbook.setCellInput(sheetId, [1, 9], '=A10*2')
    workbook.setNumberFormat(sheetId, { x: 1, y: 9, width: 1, height: 1 }, 'currency')

    workbook.fillRange(
      sheetId,
      { x: 1, y: 9, width: 1, height: 1 },
      { x: 1, y: 10, width: 1, height: 1 },
    )

    expect(workbook.getCellInput(sheetId, [1, 10])).toBe('=A11*2')
    expect(workbook.getCellValue(sheetId, [1, 10])).toBe(6)
    expect(workbook.getStyle(sheetId, [1, 10]).numberFormat).toBe('currency')
  })

  it('finds, replaces, and resolves named cell ranges', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id
    if (sheetId === undefined) throw new Error('Starter sheet is missing.')

    expect(workbook.find('north')).toHaveLength(1)
    expect(workbook.replaceAll('North', 'Central')).toBe(1)
    expect(workbook.getCellValue(sheetId, [0, 1])).toBe('Central')
    expect(workbook.addNamedRange('TotalRevenue', sheetId, { x: 1, y: 5, width: 1, height: 1 })).toBe(true)
    workbook.setCellInput(sheetId, [5, 0], '=TotalRevenue+1')

    expect(workbook.getCellValue(sheetId, [5, 0])).toBe(144_111)
  })

  it('is clean again after undoing back to the saved state, and stays dirty once that state is unreachable', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    expect(workbook.isDirty()).toBe(false)

    workbook.setCellInput(sheetId, [1, 1], 1)
    workbook.markSaved()
    workbook.setCellInput(sheetId, [1, 1], 2)
    expect(workbook.isDirty()).toBe(true)
    workbook.undo()
    expect(workbook.isDirty()).toBe(false)

    // Undo past the save, then branch: the saved state now lives only on the discarded redo branch.
    workbook.undo()
    workbook.setCellInput(sheetId, [1, 1], 3)
    workbook.undo()
    expect(workbook.isDirty()).toBe(true)
  })

  it('stays dirty when an edit lands while a save is being written', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    workbook.setCellInput(sheetId, [1, 1], 1)
    const point = workbook.savePoint()
    workbook.setCellInput(sheetId, [1, 1], 2)
    workbook.markSaved(point)
    expect(workbook.isDirty()).toBe(true)
  })

  it('replaces text literally, including $ sequences, with and without match case', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    workbook.replaceAll('North', '$&bar', true)
    workbook.replaceAll('SOUTH', '$1', false)
    expect(workbook.getCellInput(sheetId, [0, 1])).toBe('$&bar')
    expect(workbook.getCellInput(sheetId, [0, 2])).toBe('$1')
  })

  it('sorts whole rows so each row keeps its own formula results and totals stay correct', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    const northMargin = workbook.getCellValue(sheetId, [3, 1])

    workbook.sortRows(sheetId, { x: 0, y: 1, width: 4, height: 4 }, 1, true)

    expect([1, 2, 3, 4].map(row => workbook.getCellValue(sheetId, [0, row]))).toEqual(['East', 'West', 'South', 'North'])
    expect(workbook.getCellValue(sheetId, [3, 4])).toBe(northMargin)
    expect(workbook.getCellValue(sheetId, [1, 5])).toBe(144_110)
  })

  it('treats a text header row as fixed when sorting from a single cell', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    expect(workbook.sortRegion(sheetId)).toEqual({ x: 0, y: 1, width: 4, height: 5 })
  })

  it('sums the run of numbers directly above with AutoSum', () => {
    const workbook = new Workbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    workbook.setBlock(sheetId, [0, 0], [['Item'], ['4'], ['6']])

    expect(workbook.autoSum(sheetId, [0, 3])).toBe(true)
    expect(workbook.getCellInput(sheetId, [0, 3])).toBe('=SUM(A2:A3)')
    expect(workbook.getCellValue(sheetId, [0, 3])).toBe(10)
  })

  it('clears a range in one undo step and keeps its formatting', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    workbook.setNumberFormat(sheetId, { x: 1, y: 1, width: 1, height: 1 }, 'currency')

    workbook.clearContents(sheetId, { x: 1, y: 1, width: 2, height: 4 })
    expect(workbook.getCellValue(sheetId, [1, 1])).toBeNull()
    expect(workbook.getCellValue(sheetId, [1, 5])).toBe(0)
    expect(workbook.getStyle(sheetId, [1, 1]).numberFormat).toBe('currency')

    workbook.undo()
    expect(workbook.getCellValue(sheetId, [1, 5])).toBe(144_110)
  })
})
