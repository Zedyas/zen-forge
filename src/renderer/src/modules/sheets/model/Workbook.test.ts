import { describe, expect, it } from 'vitest'
import { fillLines } from '../grid-mapping'
import { sheetNameFrom, sheetNameProblem } from './address'
import { blankWorkbookData, Workbook } from './Workbook'
import type { CellStyle, WorkbookData } from './workbook-data'

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

  it('reads TRUE and FALSE in formulas as Excel does, and keeps them out of the named ranges', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0

    workbook.setCellInput(sheetId, [5, 0], '=VLOOKUP("South",A2:B5,2,FALSE)')
    workbook.setCellInput(sheetId, [5, 1], '=IF(B2>0,TRUE,false)')

    expect(workbook.getCellValue(sheetId, [5, 0])).toBe(38_750)
    expect(workbook.getCellValue(sheetId, [5, 1])).toBe(true)
    expect(workbook.getCellInput(sheetId, [5, 0])).toBe('=VLOOKUP("South",A2:B5,2,FALSE)')
    expect(workbook.namedRanges()).toEqual([])
    expect(workbook.toData().namedRanges).toEqual([])
  })

  it('fits a row to wrapped text on a sheet with more styled cells than a function call can take arguments', () => {
    const blank = blankWorkbookData()
    const sheet = blank.sheets[0]
    if (sheet === undefined) throw new Error('Blank workbook has no sheet.')
    const wrapped: CellStyle = { numberFormat: 'general', decimalPlaces: 2, wrap: true }
    const styles = Array.from({ length: 130_000 }, (_, index) => [index % 5, Math.floor(index / 5), wrapped] as const)
    const workbook = new Workbook({ ...blank, sheets: [{ ...sheet, styles }] })
    const sheetId = workbook.sheets()[0]?.id ?? 0

    workbook.setCellInput(sheetId, [0, 0], 'long wrapped text', () => 60)

    expect(workbook.getRowHeight(sheetId, 0)).toBe(60)
  })

  it('holds as many rows as Excel, beyond the engine’s default limit of 40,000', () => {
    const blank = blankWorkbookData()
    const sheet = blank.sheets[0]
    if (sheet === undefined) throw new Error('Blank workbook has no sheet.')
    const cells = Array.from({ length: 50_000 }, (_, row) => [row + 1])
    const workbook = new Workbook({ ...blank, sheets: [{ ...sheet, cells }] })
    const sheetId = workbook.sheets()[0]?.id ?? 0

    workbook.setCellInput(sheetId, [1, 0], '=SUM(A:A)')

    expect(workbook.getCellValue(sheetId, [0, 49_999])).toBe(50_000)
    expect(workbook.getCellValue(sheetId, [1, 0])).toBe(1_250_025_000)
  })

  it('adds a sheet under a free name when the next number is taken', () => {
    const blank = blankWorkbookData()
    const sheet = blank.sheets[0]
    if (sheet === undefined) throw new Error('Blank workbook has no sheet.')
    const workbook = new Workbook({ ...blank, sheets: [sheet, { ...sheet, name: 'Sheet3' }] })

    workbook.addSheet()
    const renamed = workbook.sheets()[0]?.id ?? 0
    workbook.renameSheet(renamed, 'sheet5')
    workbook.addSheet()

    expect(workbook.sheets().map(candidate => candidate.name)).toEqual(['sheet5', 'Sheet3', 'Sheet4', 'Sheet6'])
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

    workbook.setNumberFormat(sheetId, [{ x: 0, y: 0, width: 1, height: 1 }], 'currency')
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
    workbook.setNumberFormat(sheetId, [{ x: 1, y: 9, width: 1, height: 1 }], 'currency')

    workbook.fillCells(sheetId, [{ seeds: [[1, 9]], targets: [[1, 10]] }], false)

    expect(workbook.getCellInput(sheetId, [1, 10])).toBe('=A11*2')
    expect(workbook.getCellValue(sheetId, [1, 10])).toBe(6)
    expect(workbook.getStyle(sheetId, [1, 10]).numberFormat).toBe('currency')
  })

  it('continues a series past a hidden row and fills up from the seeds, each as one undo step', () => {
    const workbook = new Workbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    workbook.setBlock(sheetId, [0, 5], [['1'], ['2']])
    workbook.hideRows(sheetId, 8)
    const { lines, backward } = fillLines(
      { x: 0, y: 5, width: 1, height: 2 },
      { x: 0, y: 5, width: 1, height: 4 },
      [0],
      [0, 1, 2, 3, 4, 5, 6, 7, 9, 10],
    )

    workbook.fillCells(sheetId, lines, backward)
    workbook.fillCells(sheetId, [{ seeds: [[0, 5], [0, 6]], targets: [[0, 4], [0, 3]] }], true)

    expect([3, 4, 7, 8, 9].map(row => workbook.getCellValue(sheetId, [0, row]))).toEqual([-1, 0, 3, null, 4])
    workbook.undo()
    expect(workbook.getCellValue(sheetId, [0, 4])).toBeNull()
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

  it('pastes 150,000 cells in well under the test timeout, leaving a skipped row as it was', () => {
    const workbook = new Workbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    workbook.setCellInput(sheetId, [0, 500], 'kept')
    // Row 500 is left out, as a hidden row is: the rows around it still go in as blocks.
    const cells = Array.from({ length: 75_000 }, (_, row) => row).filter(row => row !== 500)
      .flatMap(row => [{ target: [0, row] as const, input: row }, { target: [1, row] as const, input: `=A${row + 1}*2` }])
    workbook.pasteCells(sheetId, cells)

    expect(workbook.getCellValue(sheetId, [0, 74_999])).toBe(74_999)
    expect(workbook.getCellValue(sheetId, [1, 74_999])).toBe(149_998)
    expect(workbook.getCellValue(sheetId, [0, 500])).toBe('kept')
  })

  it('never writes the cells a merged cell covers', () => {
    const workbook = new Workbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    workbook.setCellInput(sheetId, [0, 0], 5)
    workbook.mergeCells(sheetId, { x: 0, y: 1, width: 1, height: 2 }, false)

    expect(workbook.setCellInput(sheetId, [0, 2], 7)).toBe(false)
    expect(workbook.fillCells(sheetId, [{ seeds: [[0, 0]], targets: [[0, 1], [0, 2]] }], false)).toBe(false)
    workbook.pasteCells(sheetId, [{ target: [0, 1], input: 1 }, { target: [0, 2], input: 2 }])
    expect([1, 2].map(row => workbook.getCellValue(sheetId, [0, row]))).toEqual([1, null])
  })

  it('moves a named range with its cut cells, through undo, redo and a reopen', () => {
    const workbook = new Workbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    workbook.setCellInput(sheetId, [0, 0], 7)
    workbook.addNamedRange('Foo', sheetId, { x: 0, y: 0, width: 1, height: 1 })
    workbook.setCellInput(sheetId, [1, 0], '=Foo')

    workbook.moveCells(sheetId, { x: 0, y: 0, width: 1, height: 1 }, sheetId, [0, 4])
    expect(workbook.namedRanges()[0]?.range).toEqual({ x: 0, y: 4, width: 1, height: 1 })
    expect(new Workbook(workbook.toData()).getCellValue(sheetId, [1, 0])).toBe(7)
    workbook.undo()
    expect(workbook.namedRanges()[0]?.range).toEqual({ x: 0, y: 0, width: 1, height: 1 })
    workbook.redo()
    expect(workbook.getCellValue(sheetId, [1, 0])).toBe(7)
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

  it('formats and clears every area of a selection that skips hidden rows, as one undo step', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    workbook.hideRows(sheetId, 2)
    const areas = [{ x: 1, y: 1, width: 1, height: 1 }, { x: 1, y: 3, width: 1, height: 1 }]

    workbook.updateStyle(sheetId, areas, { bold: true })
    workbook.clearContents(sheetId, areas)

    expect(workbook.getStyle(sheetId, [1, 2]).bold).toBeUndefined()
    expect([1, 2, 3].map(row => workbook.getCellValue(sheetId, [1, row]))).toEqual([null, 38_750, null])
    expect(workbook.selectionStats(sheetId, areas).count).toBe(0)
    workbook.undo()
    expect(workbook.getCellValue(sheetId, [1, 3])).toBe(29_400)
  })

  it('keeps a format code set with the number format, and drops it once the format changes', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    const areas = [{ x: 1, y: 1, width: 1, height: 1 }]

    workbook.updateStyle(sheetId, areas, { numberFormat: 'general', decimalPlaces: 2, formatCode: '0.00E+00' })
    expect(workbook.getStyle(sheetId, [1, 1]).formatCode).toBe('0.00E+00')
    workbook.changeDecimalPlaces(sheetId, areas, 1)
    expect(workbook.getStyle(sheetId, [1, 1])).toEqual({ numberFormat: 'number', decimalPlaces: 3 })
  })

  it('renames sheets only to names Excel accepts', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    workbook.addSheet()

    expect(sheetNameProblem('Q1 [draft]', [])).toBe('A sheet name can’t contain / \\ ? * [ ] or :')
    expect(sheetNameProblem("'Totals", [])).toBe('A sheet name can’t start or end with an apostrophe.')
    expect(sheetNameProblem('x'.repeat(32), [])).toBe('A sheet name can be at most 31 characters.')
    expect(workbook.renameSheet(sheetId, 'sheet2')).toBe(false)
    expect(workbook.renameSheet(sheetId, 'a/b')).toBe(false)
    expect(workbook.renameSheet(sheetId, 'sheet1')).toBe(true)
    expect(workbook.sheets().map(sheet => sheet.name)).toEqual(['sheet1', 'Sheet2'])
    expect(sheetNameFrom("Budget: 2026 [final]'")).toBe('Budget_ 2026 _final_')
  })

  it('fits rows to wrapped text in the same undo step as the wrap, and moves heights with inserted rows', () => {
    const workbook = new Workbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    // Wrapped text needs 20 px per 10 characters; anything else fits the default row.
    const measure = (text: string, style: CellStyle): number => style.wrap === true ? Math.ceil(text.length / 10) * 20 : 26
    workbook.setCellInput(sheetId, [0, 1], 'forty characters of text that need wrap')

    workbook.updateStyle(sheetId, [{ x: 0, y: 1, width: 1, height: 1 }], { wrap: true }, measure)
    expect(workbook.getRowHeight(sheetId, 1)).toBe(80)
    workbook.insertRows(sheetId, 0)
    expect(workbook.getRowHeight(sheetId, 2)).toBe(80)
    workbook.undo()
    workbook.undo()
    expect(workbook.getRowHeight(sheetId, 1)).toBe(26)

    workbook.setRowHeight(sheetId, [3, 4], 40)
    workbook.fitRows(sheetId, [3], measure)
    expect([3, 4].map(row => workbook.getRowHeight(sheetId, row))).toEqual([26, 40])
  })

  it('merges keeping only the top-left value, and keeps merges consistent through inserts, deletes and sorts', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    const block = { x: 0, y: 1, width: 2, height: 2 }

    expect(workbook.mergeDiscards(sheetId, block, false)).toBe(3)
    expect(workbook.mergeDiscards(sheetId, block, true)).toBe(2)
    workbook.mergeCells(sheetId, block, false)
    expect(workbook.getCellValue(sheetId, [0, 1])).toBe('North')
    expect([[1, 1], [0, 2], [1, 2]].map(([col = 0, row = 0]) => workbook.getCellValue(sheetId, [col, row]))).toEqual([null, null, null])
    expect(workbook.mergeAt(sheetId, [1, 2])).toEqual(block)

    workbook.insertRows(sheetId, 2)
    expect(workbook.merges(sheetId)).toEqual([{ x: 0, y: 1, width: 2, height: 3 }])
    workbook.deleteRows(sheetId, 0, 2)
    expect(workbook.merges(sheetId)).toEqual([{ x: 0, y: 0, width: 2, height: 2 }])
    expect(workbook.sortRows(sheetId, { x: 0, y: 0, width: 4, height: 5 }, 1, true)).toBe('merged')

    workbook.unmergeCells(sheetId, { x: 1, y: 1, width: 1, height: 1 })
    expect(workbook.merges(sheetId)).toEqual([])
    workbook.undo()
    workbook.undo()
    workbook.undo()
    workbook.undo()
    expect(workbook.merges(sheetId)).toEqual([])
    expect(workbook.getCellValue(sheetId, [1, 1])).toBe(42_100)
  })

  it('clears a range in one undo step and keeps its formatting', () => {
    const workbook = regionsWorkbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    workbook.setNumberFormat(sheetId, [{ x: 1, y: 1, width: 1, height: 1 }], 'currency')

    workbook.clearContents(sheetId, [{ x: 1, y: 1, width: 2, height: 4 }])
    expect(workbook.getCellValue(sheetId, [1, 1])).toBeNull()
    expect(workbook.getCellValue(sheetId, [1, 5])).toBe(0)
    expect(workbook.getStyle(sheetId, [1, 1]).numberFormat).toBe('currency')

    workbook.undo()
    expect(workbook.getCellValue(sheetId, [1, 5])).toBe(144_110)
  })
})
