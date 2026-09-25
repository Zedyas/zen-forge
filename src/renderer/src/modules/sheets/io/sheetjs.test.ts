import * as XLSX from 'xlsx'
import numbersTemplate from 'xlsx/dist/xlsx.zahl'
import { describe, expect, it } from 'vitest'
import { readSheetJs } from './sheetjs'

/** 2024-01-15 as a 1900-system serial, the only date system the workbook model uses. */
const dueDate = 45_306

/** Written with SheetJS itself, so each test reads what the format's own writer produces. */
function fixture(bookType: XLSX.BookType): Uint8Array {
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Item', 'Cost', 'Paid', 'Due'],
    ['Rent', 2150.5, true],
    ['Food', 612.5, false],
  ])
  // SheetJS's Numbers writer takes serials in the 1904 date system, which Numbers files use.
  sheet['D2'] = { t: 'n', v: bookType === 'numbers' ? dueDate - 1_462 : dueDate, z: 'm/d/yy' }
  sheet['B4'] = { t: 'n', f: 'SUM(B2:B3)', v: 2763 }
  sheet['A5'] = { t: 's', v: 'Spans two columns' }
  sheet['!ref'] = 'A1:D5'
  sheet['!merges'] = [XLSX.utils.decode_range('A5:B5')]
  sheet['!cols'] = [{ wch: 20 }, {}, {}, { hidden: true }]
  const bytes: ArrayBuffer = XLSX.write(XLSX.utils.book_new(sheet, 'Costs'), { bookType, type: 'array', numbers: numbersTemplate })
  return new Uint8Array(bytes)
}

async function read(bookType: XLSX.BookType, format: 'xls' | 'ods' | 'numbers') {
  const { data, findings } = await readSheetJs(fixture(bookType), format)
  const sheet = data.sheets[0]
  return { sheet, constructs: findings.map(finding => finding.construct) }
}

describe('SheetJS formats', () => {
  it('reads an Excel 97–2003 workbook', async () => {
    const { sheet, constructs } = await read('biff8', 'xls')

    expect(sheet?.cells.slice(0, 3)).toEqual([
      ['Item', 'Cost', 'Paid', 'Due'],
      ['Rent', 2150.5, true, dueDate],
      ['Food', 612.5, false],
    ])
    // SheetJS writes xls formulas as their values only.
    expect(sheet?.values?.[3]?.[1]).toBe(2763)
    expect(sheet?.styles).toContainEqual([3, 1, expect.objectContaining({ numberFormat: 'date' })])
    // Nor does it write hidden rows to xls; a hidden column exercises the same layout path.
    expect(sheet?.layout.hiddenColumns).toEqual([3])
    expect(sheet?.layout.columnWidths).toContainEqual([0, expect.any(Number)])
    expect(constructs).toEqual(expect.arrayContaining(['Merged cells', 'Cell formatting']))
  })

  it('reads an OpenDocument spreadsheet', async () => {
    const { sheet, constructs } = await read('ods', 'ods')

    expect(sheet?.cells[1]).toEqual(['Rent', 2150.5, true, dueDate])
    expect(sheet?.cells[3]?.[1]).toBe('=SUM(B2:B3)')
    expect(sheet?.values?.[3]?.[1]).toBe(2763)
    expect(sheet?.styles).toContainEqual([3, 1, expect.objectContaining({ numberFormat: 'date' })])
    // The ods reader ignores hidden rows and columns, so the report has to say so.
    expect(sheet?.layout.hiddenColumns).toEqual([])
    expect(constructs).toEqual(expect.arrayContaining([
      'Merged cells',
      'Cell formatting',
      'Column widths, hidden rows and columns, and frozen panes',
    ]))
  })

  it('reads a Numbers document, moving its dates to the 1900 system', async () => {
    const { sheet, constructs } = await read('numbers', 'numbers')

    expect(sheet?.cells[1]).toEqual(['Rent', 2150.5, true, dueDate])
    // Numbers formulas arrive as their last calculated values.
    expect(sheet?.cells[3]?.[1]).toBe(2763)
    expect(sheet?.styles).toContainEqual([3, 1, expect.objectContaining({ numberFormat: 'date' })])
    expect(constructs).toEqual(expect.arrayContaining([
      'Merged cells',
      'Formulas',
      'Column widths, hidden rows and columns, and frozen panes',
    ]))
  })
})
