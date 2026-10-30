import { describe, expect, it } from 'vitest'
import { textCellInput } from '../model/cell-parse'
import { delimitedValue } from '../model/format'
import { Workbook } from '../model/Workbook'
import { readDelimited, writeDelimited } from './csv'

/** windows-1252 is a single-byte encoding, so each character maps to its own code point. */
function encodeWindows1252(text: string): Uint8Array {
  return Uint8Array.from([...text], character => character.charCodeAt(0))
}

describe('csv', () => {
  it('reads a semicolon-delimited windows-1252 file with quoted newlines and typed numbers', () => {
    const bytes = encodeWindows1252(
      'Name;Amount;Note\nJosé;$1,234.50;"line one,\nline two"\n',
    )

    const { data, findings } = readDelimited(bytes, 'Imported', 'csv')
    const sheet = data.sheets[0]

    expect(sheet?.cells).toEqual([
      ["'Name", "'Amount", "'Note"],
      ["'José", 1234.5, "'line one,\nline two"],
    ])
    expect(sheet?.styles).toEqual([[1, 1, { numberFormat: 'currency', decimalPlaces: 2 }]])
    expect(findings.map(finding => finding.construct)).toEqual(['Non-Unicode text encoding'])
  })

  it('keeps IDs, ZIP codes and badly grouped numbers as text, and reads valid grouping as numbers', () => {
    const bytes = new TextEncoder().encode('ID,Zip,Card,Amount,Odd\nA1,02134,4111111111111111,"1,234.50","1,5"\n')

    const cells = readDelimited(bytes, 'Imported', 'csv').data.sheets[0]?.cells

    expect(cells?.[1]).toEqual(["'A1", "'02134", "'4111111111111111", 1234.5, "'1,5"])
  })

  it('writes underlying values: full precision, ISO dates and formula results', () => {
    const workbook = new Workbook()
    const sheetId = workbook.sheets()[0]?.id ?? 0
    workbook.setCellInput(sheetId, [0, 0], 1234.5678)
    workbook.setCellInput(sheetId, [1, 0], 46_247.5)
    workbook.setCellInput(sheetId, [2, 0], '=A1*2')
    workbook.setCellInput(sheetId, [3, 0], textCellInput('00123'))
    workbook.setNumberFormat(sheetId, [{ x: 0, y: 0, width: 1, height: 1 }], 'number')
    workbook.setNumberFormat(sheetId, [{ x: 1, y: 0, width: 1, height: 1 }], 'date')

    const row = [0, 1, 2, 3].map(col => delimitedValue(workbook.getCellValue(sheetId, [col, 0]), workbook.getStyle(sheetId, [col, 0])))

    // The grid shows 1,234.57 and a locale date; the file keeps the values themselves.
    expect(row).toEqual(['1234.5678', '2026-08-13 12:00:00', '2469.1356', '00123'])
    const back = readDelimited(writeDelimited([row], 'csv'), 'Back', 'csv').data.sheets[0]
    expect(back?.cells[0]).toEqual([1234.5678, 46_247.5, 2469.1356, "'00123"])
    expect(back?.styles).toContainEqual([1, 0, { numberFormat: 'date', decimalPlaces: 0 }])
  })

  it('quotes separators, quotes and newlines and writes a UTF-8 BOM', () => {
    const bytes = writeDelimited([
      ['plain', 'has,comma', 'has"quote'],
      ['has\nnewline', 'é', ''],
    ], 'csv')

    expect([...bytes.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf])
    expect(new TextDecoder().decode(bytes.subarray(3))).toBe(
      'plain,"has,comma","has""quote"\r\n"has\nnewline",é,',
    )
  })
})

describe('tsv', () => {
  it('round-trips tabs, quotes and newlines inside fields and leaves commas unquoted', () => {
    const rows = [
      ['Item', 'Note', 'Cost'],
      ['Desk, oak', 'left\tright', '120'],
      ['Lamp', 'say "hi"\non two lines', '35.5'],
    ]

    const bytes = writeDelimited(rows, 'tsv')
    expect(new TextDecoder().decode(bytes.subarray(3))).toBe(
      'Item\tNote\tCost\r\nDesk, oak\t"left\tright"\t120\r\nLamp\t"say ""hi""\non two lines"\t35.5',
    )

    const { data, findings } = readDelimited(bytes, 'Imported', 'tsv')
    expect(data.sheets[0]?.cells).toEqual([
      ["'Item", "'Note", "'Cost"],
      ["'Desk, oak", "'left\tright", 120],
      ["'Lamp", "'say \"hi\"\non two lines", 35.5],
    ])
    expect(findings).toEqual([])
  })
})
