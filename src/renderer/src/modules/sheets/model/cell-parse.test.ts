import { describe, expect, it } from 'vitest'
import { parseCellText, textCellInput } from './cell-parse'
import { Workbook } from './Workbook'

describe('parseCellText', () => {
  it.each([
    ['0', 0],
    ['0.5', 0.5],
    ['-12.25', -12.25],
    ['1,234', 1234],
    ['12,345,678.5', 12_345_678.5],
    ['$1234.50', 1234.5],
    ['1e+21', 1e21],
    ['123456789012345', 123_456_789_012_345],
    ['2026-08-13 12:00:00', 46_247.5],
    ['2026-08-13T06:00', 46_247.25],
    ['2026-08-13', 46_247],
  ])('reads %s as a number', (text, number) => {
    expect(parseCellText(text).input).toBe(number)
  })

  it.each([
    ['00123', "'00123"],
    ['-007', "'-007"],
    ['1234567890123456', "'1234567890123456"],
    ['4111 1111 1111 1111', "'4111 1111 1111 1111"],
    ['1,5', "'1,5"],
    ['12,34', "'12,34"],
    ['$1,5', "'$1,5"],
    ['12E3', "'12E3"],
    ['2026-02-30 10:00', "'2026-02-30 10:00"],
  ])('keeps %s as text', (text, input) => {
    expect(parseCellText(text).input).toBe(input)
  })

  it('reads an empty field as an empty cell, not empty text', () => {
    expect(parseCellText(' ').input).toBeNull()
  })
})

describe('textCellInput', () => {
  it('keeps every text as text, and shows the mark only where typing the text alone would read as a value', () => {
    const workbook = new Workbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    const texts = ['00123', '=x', 'TRUE', '#N/A', '10:30', '$-12', 'plain words', "'quoted"]
    texts.forEach((text, row) => workbook.setCellInput(sheetId, [0, row], textCellInput(text)))

    expect(texts.map((_, row) => workbook.getCellValue(sheetId, [0, row]))).toEqual(texts)
    expect(texts.map((_, row) => workbook.getCellInput(sheetId, [0, row])))
      .toEqual(["'00123", "'=x", "'TRUE", "'#N/A", "'10:30", "'$-12", 'plain words', "''quoted"])
  })
})
