import { describe, expect, it } from 'vitest'
import { a4Paper, letterPaper } from '@shared/print'
import { blankWorkbookData, Workbook } from './model/Workbook'
import { planSheetPages, sheetPrintTable, type PrintCell } from './print'

function workbookWith(cells: (string | number)[][]): { workbook: Workbook; sheetId: number } {
  const blank = blankWorkbookData()
  const sheet = blank.sheets[0]
  if (sheet === undefined) throw new Error('Blank workbook has no sheet.')
  const workbook = new Workbook({ ...blank, sheets: [{ ...sheet, cells }] })
  const sheetId = workbook.sheets()[0]?.id
  if (sheetId === undefined) throw new Error('Workbook has no sheet.')
  return { workbook, sheetId }
}

const texts = (rows: readonly (readonly PrintCell[])[]): string[][] => rows.map(row => row.map(cell => cell.text))

describe('sheetPrintTable', () => {
  it('leaves out hidden rows and columns and prints values as the grid shows them', () => {
    const { workbook, sheetId } = workbookWith([
      ['Item', 'Note', 'Share'],
      ['Rent', 'secret', 0.375],
      ['Food', 'hidden row', 0.125],
      ['Total', '', '=C2+C3'],
    ])
    workbook.hideColumns(sheetId, 1)
    workbook.hideRows(sheetId, 2)
    workbook.setNumberFormat(sheetId, { x: 2, y: 1, width: 1, height: 3 }, 'percent')

    const table = sheetPrintTable(workbook, sheetId)

    expect(table?.columnWidths).toHaveLength(2)
    expect(texts(table?.bodyRows ?? [])).toEqual([
      ['Item', 'Share'],
      ['Rent', expect.stringMatching(/^37[.,]50\s?%$/)],
      ['Total', expect.stringMatching(/^50[.,]00\s?%$/)],
    ])
    expect(table?.bodyRows[1]?.[1]?.align).toBe('right')
  })

  it('puts the visible frozen rows in the header that repeats on every page', () => {
    const { workbook, sheetId } = workbookWith([['Title'], ['Month'], ['Jan'], ['Feb']])
    workbook.freezeRows(sheetId, 2)
    workbook.hideRows(sheetId, 0)

    const table = sheetPrintTable(workbook, sheetId)

    expect(texts(table?.headRows ?? [])).toEqual([['Month']])
    expect(texts(table?.bodyRows ?? [])).toEqual([['Jan'], ['Feb']])
  })

  it('has nothing to print when the visible cells are empty', () => {
    const empty = workbookWith([])
    expect(sheetPrintTable(empty.workbook, empty.sheetId)).toBeUndefined()
    const hidden = workbookWith([['only in a hidden column']])
    hidden.workbook.hideColumns(hidden.sheetId, 0)
    expect(sheetPrintTable(hidden.workbook, hidden.sheetId)).toBeUndefined()
  })
})

describe('planSheetPages', () => {
  it('keeps a table that fits a portrait page at full size', () => {
    const plan = planSheetPages([120, 120, 120, 120, 120], 0, letterPaper)
    expect(plan.page).toMatchObject({ width: letterPaper.width, height: letterPaper.height })
    expect(plan.scale).toBe(1)
    expect(plan.columnGroups).toEqual([[0, 1, 2, 3, 4]])
  })

  it('turns the page for a clearly wider table and scales it to the page width', () => {
    const plan = planSheetPages(Array.from({ length: 10 }, () => 120), 0, a4Paper)
    expect(plan.page).toMatchObject({ width: a4Paper.height, height: a4Paper.width })
    expect(plan.scale).toBeGreaterThan(0.6)
    expect(plan.scale).toBeLessThan(1)
    expect(plan.columnGroups).toHaveLength(1)
  })

  it('continues a very wide table on more pages across, repeating the frozen columns', () => {
    const plan = planSheetPages(Array.from({ length: 30 }, () => 120), 1, a4Paper)
    expect(plan.scale).toBe(0.6)
    expect(plan.columnGroups.length).toBeGreaterThan(1)
    expect(plan.columnGroups.every(group => group[0] === 0)).toBe(true)
    expect(plan.columnGroups.flatMap(group => group.slice(1))).toEqual(Array.from({ length: 29 }, (_, index) => index + 1))
  })
})
