// @vitest-environment happy-dom
import { describe, expect, it } from 'vitest'
import { copyCells, pasteCells, pasteTargets } from './clipboard'
import { Workbook } from './model/Workbook'

const all = Array.from({ length: 30 }, (_, index) => index)

function sheetWith(cells: readonly (readonly (string | number)[])[]): { workbook: Workbook; sheetId: number } {
  const workbook = new Workbook()
  const sheetId = workbook.sheets()[0]?.id ?? 0
  cells.forEach((row, y) => row.forEach((input, x) => workbook.setCellInput(sheetId, [x, y], input)))
  return { workbook, sheetId }
}

describe('Ledger clipboard', () => {
  it('pastes formulas with their references moved, styles, and the shown values as text', () => {
    const { workbook, sheetId } = sheetWith([[1, 2, '=A1+B1', '=$A$1*2']])
    workbook.updateStyle(sheetId, [{ x: 2, y: 0, width: 1, height: 1 }], { bold: true, numberFormat: 'currency' })
    const data = new DataTransfer()

    copyCells(data, { workbook, sheetId, columns: [2, 3], rows: [0], range: { x: 2, y: 0, width: 2, height: 1 } }, false)
    pasteCells(data, { workbook, sheetId, selection: { x: 2, y: 3, width: 1, height: 1 }, columns: all, rows: all })

    expect(data.getData('text/plain')).toMatch(/^\$3\.00\t2$/)
    expect(workbook.getCellInput(sheetId, [2, 3])).toBe('=A4+B4')
    expect(workbook.getCellInput(sheetId, [3, 3])).toBe('=$A$1*2')
    expect(workbook.getStyle(sheetId, [2, 3])).toMatchObject({ bold: true, numberFormat: 'currency' })
  })

  it('moves cut cells on paste, so formulas that refer to them follow, as one undo step', () => {
    const { workbook, sheetId } = sheetWith([[5, '=A1*2']])
    const data = new DataTransfer()

    copyCells(data, { workbook, sheetId, columns: [0], rows: [0], range: { x: 0, y: 0, width: 1, height: 1 } }, true)
    expect(workbook.getCellValue(sheetId, [0, 0])).toBe(5)
    pasteCells(data, { workbook, sheetId, selection: { x: 0, y: 4, width: 1, height: 1 }, columns: all, rows: all })

    expect(workbook.getCellValue(sheetId, [0, 0])).toBeNull()
    expect(workbook.getCellValue(sheetId, [0, 4])).toBe(5)
    expect(workbook.getCellInput(sheetId, [1, 0])).toBe('=A5*2')
    workbook.undo()
    expect(workbook.getCellValue(sheetId, [0, 0])).toBe(5)
    expect(workbook.getCellInput(sheetId, [1, 0])).toBe('=A1*2')
  })

  it('pastes tab-separated text from other apps as typed values', () => {
    const { workbook, sheetId } = sheetWith([])
    const data = new DataTransfer()
    data.setData('text/plain', 'Name\tAmount\r\n"Desk, oak"\t$1,250.00\r\n')

    pasteCells(data, { workbook, sheetId, selection: { x: 1, y: 1, width: 1, height: 1 }, columns: all, rows: all })

    expect(workbook.getCellValue(sheetId, [1, 2])).toBe('Desk, oak')
    expect(workbook.getCellValue(sheetId, [2, 2])).toBe(1250)
    expect(workbook.getStyle(sheetId, [2, 2]).numberFormat).toBe('currency')
  })

  it('places a block on visible rows and repeats it over a selection that is a multiple of it', () => {
    // Row 1 is hidden.
    const rows = [0, 2, 3, 4]
    expect(pasteTargets(1, 1, { x: 0, y: 0, width: 1, height: 3 }, [0], rows).map(cell => cell.target)).toEqual([[0, 0], [0, 2], [0, 3]])
    expect(pasteTargets(2, 1, { x: 0, y: 0, width: 1, height: 3 }, [0], rows).map(cell => cell.target)).toEqual([[0, 0], [0, 2]])
  })
})
