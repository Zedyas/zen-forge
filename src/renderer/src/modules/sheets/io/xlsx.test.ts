import * as ExcelJS from 'exceljs'
import { unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { formatCellValue } from '../model/format'
import { Workbook } from '../model/Workbook'
import { defaultCellStyle, type WorkbookData } from '../model/workbook-data'
import { readXlsx, writeXlsx } from './xlsx'

/** A workbook that exercises every construct the format layer claims to preserve. */
const workbook: WorkbookData = {
  sheets: [
    {
      name: 'Sales',
      cells: [
        ["'Region", "'Revenue", "'Share", "'Booked"],
        ["'North", 42_100, '=B2/Totals!B1', 45_000],
        ["'South", 38_750, '=B3/Totals!B1', 45_001],
        [null, '=SUM(B2:B3)'],
      ],
      values: [
        ['Region', 'Revenue', 'Share', 'Booked'],
        ['North', 42_100, 0.9355555555555556, 45_000],
        ['South', 38_750, 0.8611111111111112, 45_001],
        [null, 80_850],
      ],
      styles: [
        [0, 0, { numberFormat: 'general', decimalPlaces: 2, bold: true, fillColor: '#eef2ff', align: 'center' }],
        [1, 0, { numberFormat: 'general', decimalPlaces: 2, strike: true, wrap: true, fontFamily: 'Georgia', fontSize: 14 }],
        [0, 1, {
          numberFormat: 'general',
          decimalPlaces: 2,
          italic: true,
          underline: true,
          textColor: '#b91c1c',
          link: 'https://example.com/north',
        }],
        [1, 1, { numberFormat: 'currency', decimalPlaces: 2, note: 'Includes refunds.\nSee Totals.' }],
        [2, 1, { numberFormat: 'percent', decimalPlaces: 1 }],
        [3, 1, { numberFormat: 'date', decimalPlaces: 0 }],
        [1, 2, { numberFormat: 'currency', decimalPlaces: 2 }],
        [2, 2, { numberFormat: 'percent', decimalPlaces: 1 }],
        [3, 2, { numberFormat: 'date', decimalPlaces: 0 }],
        [1, 3, { numberFormat: 'currency', decimalPlaces: 0, bold: true }],
      ],
      layout: {
        rowCount: 1_000,
        columnCount: 26,
        columnWidths: [[0, 160], [2, 90]],
        rowHeights: [[1, 40], [3, 18]],
        hiddenRows: [2],
        hiddenColumns: [],
        freezeRows: 1,
        freezeColumns: 1,
        merges: [{ x: 4, y: 0, width: 2, height: 2 }, { x: 0, y: 5, width: 3, height: 1 }],
      },
    },
    {
      name: 'Totals',
      cells: [["'Target", 45_000]],
      values: [['Target', 45_000]],
      styles: [[1, 0, { numberFormat: 'number', decimalPlaces: 0 }]],
      layout: {
        rowCount: 1_000,
        columnCount: 26,
        columnWidths: [],
        rowHeights: [],
        hiddenRows: [],
        hiddenColumns: [],
        freezeRows: 0,
        freezeColumns: 0,
        merges: [],
      },
    },
  ],
  namedRanges: [{ name: 'Target', sheetName: 'Totals', range: { x: 1, y: 0, width: 1, height: 1 } }],
}

describe('xlsx', () => {
  it('round trips values, formulas, styles, layout and names without findings', async () => {
    const imported = await readXlsx(await writeXlsx(workbook))

    expect(imported.findings).toEqual([])
    expect(imported.data).toEqual(workbook)
  })

  it('reports macros, pivot tables, charts and threaded comments as aggregated findings, and keeps merged cells', async () => {
    const source = new ExcelJS.Workbook()
    const sheet = source.addWorksheet('Sheet1')
    sheet.getCell('A1').value = 'merged'
    sheet.mergeCells('A1:B2')
    sheet.mergeCells('D1:E1')
    const written = new Uint8Array(await source.xlsx.writeBuffer())

    const entries = unzipSync(written)
    entries['xl/vbaProject.bin'] = new Uint8Array([1, 2, 3])
    entries['xl/pivotTables/pivotTable1.xml'] = new TextEncoder().encode('<pivotTableDefinition/>')
    entries['xl/charts/chart1.xml'] = new TextEncoder().encode('<chartSpace/>')
    entries['xl/threadedComments/threadedComment1.xml'] = new TextEncoder().encode('<ThreadedComments/>')

    const { data, findings } = await readXlsx(zipSync(entries))
    const byConstruct = new Map(findings.map(finding => [finding.construct, finding]))

    expect(byConstruct.get('Macros (VBA)')?.severity).toBe('dropped')
    expect(byConstruct.get('Pivot tables')?.severity).toBe('dropped')
    expect(byConstruct.get('Charts')?.severity).toBe('dropped')
    expect(byConstruct.get('Threaded comments')?.severity).toBe('dropped')
    expect(data.sheets[0]?.layout.merges).toEqual([{ x: 0, y: 0, width: 2, height: 2 }, { x: 3, y: 0, width: 2, height: 1 }])
    expect(byConstruct.has('Merged cells')).toBe(false)
    expect(findings.length).toBe(byConstruct.size)
    // Every note in the import report says what happens to that part of the file.
    expect(findings.filter(finding => finding.suggestedAlternative === undefined)).toEqual([])
  })

  it('translates a shared formula into each cell of its range', async () => {
    const source = new ExcelJS.Workbook()
    const sheet = source.addWorksheet('Sheet1')
    sheet.getCell('A1').value = 1
    sheet.getCell('A2').value = 2
    sheet.getCell('A3').value = 3
    sheet.fillFormula('B1:B3', 'A1*2', [2, 4, 6])

    const { data } = await readXlsx(new Uint8Array(await source.xlsx.writeBuffer()))
    const cells = data.sheets[0]?.cells

    expect(cells?.map(row => row[1])).toEqual(['=A1*2', '=A2*2', '=A3*2'])
    expect(data.sheets[0]?.values?.map(row => row[1])).toEqual([2, 4, 6])
  })

  it('calculates functions Excel saves with a _xlfn. prefix, and writes the prefix back for Excel', async () => {
    const source = new ExcelJS.Workbook()
    const sheet = source.addWorksheet('Sheet1')
    sheet.getCell('A1').value = 5
    sheet.getCell('A2').value = 7
    sheet.getCell('B1').value = { formula: '_xlfn.IFS(A1>6,"big",TRUE,"_xlfn.IFS(")', result: '_xlfn.IFS(' }
    sheet.getCell('B2').value = { formula: '_xlfn.STDEV.S(A1:A2)+_xlfn.CONCAT("a")', result: 0 }
    const opened = new Workbook((await readXlsx(new Uint8Array(await source.xlsx.writeBuffer()))).data)
    const sheetId = opened.sheets()[0]?.id ?? 0

    expect(opened.getCellInput(sheetId, [1, 0])).toBe('=IFS(A1>6,"big",TRUE,"_xlfn.IFS(")')
    expect(opened.getCellValue(sheetId, [1, 0])).toBe('_xlfn.IFS(')
    // CONCAT is not one the engine has, so it keeps its prefix, unchanged for Excel.
    expect(opened.getCellInput(sheetId, [1, 1])).toBe('=STDEV.S(A1:A2)+_xlfn.CONCAT("a")')
    opened.setCellInput(sheetId, [1, 2], '=maxifs(A1:A2,A1:A2,">0")')

    const sheetXml = new TextDecoder().decode(unzipSync(await writeXlsx(opened.toData()))['xl/worksheets/sheet1.xml'])
    const formulas = [...sheetXml.matchAll(/<f>([^<]*)<\/f>/g)].map(match => match[1]?.replaceAll('&quot;', '"').replaceAll('&gt;', '>'))
    expect(formulas).toEqual([
      '_xlfn.IFS(A1>6,"big",TRUE,"_xlfn.IFS(")',
      '_xlfn.STDEV.S(A1:A2)+_xlfn.CONCAT("a")',
      '_xlfn.MAXIFS(A1:A2,A1:A2,">0")',
    ])
  })

  it('opens a sheet with more rows than a function call can take arguments', async () => {
    const source = new ExcelJS.Workbook()
    const sheet = source.addWorksheet('Sheet1')
    for (let row = 1; row <= 130_000; row += 1) sheet.getCell(row, 1).value = row

    const { data } = await readXlsx(new Uint8Array(await source.xlsx.writeBuffer()))

    expect(data.sheets[0]?.cells).toHaveLength(130_000)
    expect(data.sheets[0]?.cells.at(-1)).toEqual([130_000])
  }, 30_000)

  it('keeps an error value an error, not text, through open and save', async () => {
    const source = new ExcelJS.Workbook()
    const sheet = source.addWorksheet('Sheet1')
    sheet.getCell('A1').value = { error: ExcelJS.ErrorValue.NotApplicable }
    sheet.getCell('B1').value = { formula: 'ISNA(A1)', result: true }
    const opened = new Workbook((await readXlsx(new Uint8Array(await source.xlsx.writeBuffer()))).data)

    const reopened = new Workbook((await readXlsx(await writeXlsx(opened.toData()))).data)
    const sheetId = reopened.sheets()[0]?.id ?? 0

    expect(reopened.getCellInput(sheetId, [0, 0])).toBe('#N/A')
    expect(reopened.getCellValue(sheetId, [1, 0])).toBe(true)
  })

  it('keeps links to places in the workbook as places, not as links to files', async () => {
    const source = new ExcelJS.Workbook()
    const sheet = source.addWorksheet('Sheet1')
    source.addWorksheet('Sheet2')
    ;['A1', 'A2', 'A3'].forEach(address => { sheet.getCell(address).value = 'go' })
    const entries = unzipSync(new Uint8Array(await source.xlsx.writeBuffer()))
    const part = 'xl/worksheets/sheet1.xml'
    entries[part] = new TextEncoder().encode(new TextDecoder().decode(entries[part]).replace('<pageMargins',
      '<hyperlinks><hyperlink ref="A1" location="Sheet2!$A$1"/><hyperlink ref="A2" location="Totals"/></hyperlinks><pageMargins'))

    const opened = new Workbook((await readXlsx(zipSync(entries))).data)
    const sheetId = opened.sheets()[0]?.id ?? 0
    opened.updateStyle(sheetId, [{ x: 0, y: 2, width: 1, height: 1 }], { link: '#Sheet2!B2' })
    expect([0, 1, 2].map(row => opened.getStyle(sheetId, [0, row]).link)).toEqual(['#Sheet2!$A$1', '#Totals', '#Sheet2!B2'])

    const saved = unzipSync(await writeXlsx(opened.toData()))
    const sheetXml = new TextDecoder().decode(saved[part])
    expect([...sheetXml.matchAll(/<hyperlink\b[^>]*>/g)].map(([tag]) => tag)).toEqual([
      '<hyperlink ref="A1" location="Sheet2!$A$1"/>',
      '<hyperlink ref="A2" location="Totals"/>',
      '<hyperlink ref="A3" location="Sheet2!B2"/>',
    ])
  })

  it('keeps every text cell text, whatever the engine would read it as, without showing the mark', async () => {
    const source = new ExcelJS.Workbook()
    const sheet = source.addWorksheet('Sheet1')
    const texts = ['$-12', '12$', '00123', 'plain']
    texts.forEach((text, row) => { sheet.getCell(row + 1, 1).value = text })
    sheet.getCell('B1').value = { formula: 'COUNT(A1:A4)' }
    const opened = new Workbook((await readXlsx(new Uint8Array(await source.xlsx.writeBuffer()))).data)
    const sheetId = opened.sheets()[0]?.id ?? 0

    expect(texts.map((_, row) => opened.getCellValue(sheetId, [0, row]))).toEqual(texts)
    expect(opened.getCellValue(sheetId, [1, 0])).toBe(0)
    // The formula bar shows the apostrophe only where typing the text alone would read as a number.
    expect(texts.map((_, row) => opened.getCellInput(sheetId, [0, row]))).toEqual(["'$-12", "'12$", "'00123", 'plain'])
    expect(opened.find('plain').map(match => match.input)).toEqual(['plain'])
    const saved = await readXlsx(await writeXlsx(opened.toData()))
    expect(saved.data.sheets[0]?.values?.map(row => row[0])).toEqual(texts)
  })

  it('keeps a row height set to Excel’s default height, and leaves rows the file did not size at Ledger’s', async () => {
    const sheet = workbook.sheets[1]
    if (sheet === undefined) throw new Error('Fixture sheet is missing.')
    const saved = await writeXlsx({ sheets: [{ ...sheet, layout: { ...sheet.layout, rowHeights: [[0, 20]] } }], namedRanges: [] })
    expect((await readXlsx(saved)).data.sheets[0]?.layout.rowHeights).toEqual([[0, 20]])

    // Excel writes `ht` on rows it sized itself, without `customHeight`.
    const entries = unzipSync(saved)
    const xml = new TextDecoder().decode(entries['xl/worksheets/sheet1.xml'])
    entries['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(xml.replace(' customHeight="1"', ''))
    expect((await readXlsx(zipSync(entries))).data.sheets[0]?.layout.rowHeights).toEqual([])
  })

  it('keeps a note on an empty cell with no other formatting', async () => {
    const sheet = workbook.sheets[1]
    if (sheet === undefined) throw new Error('Fixture sheet is missing.')
    const noted = { ...sheet, styles: [...sheet.styles, [3, 4, { ...defaultCellStyle, note: 'Fill in by Friday.' }] as const] }

    const opened = await readXlsx(await writeXlsx({ sheets: [noted], namedRanges: [] }))
    expect(opened.data.sheets[0]?.styles).toContainEqual([3, 4, { ...defaultCellStyle, note: 'Fill in by Friday.' }])
  })

  it('keeps links on numbers, booleans, formulas and empty cells without changing their values', async () => {
    const sheet = workbook.sheets[1]
    if (sheet === undefined) throw new Error('Fixture sheet is missing.')
    const targets = ['https://example.com/n', 'https://example.com/b', 'https://example.com/f?a=1&b=2', 'https://example.com/e', 'https://example.com/t']
    const linked = {
      ...sheet,
      cells: [[123, true, '=1+1', null, "'text"]],
      values: [[123, true, 2, null, 'text']],
      styles: targets.map((link, column) => [column, 0, { ...defaultCellStyle, link }] as const),
    }

    const saved = await writeXlsx({ sheets: [linked], namedRanges: [] })
    const opened = await readXlsx(saved)
    expect(opened.data.sheets[0]?.cells).toEqual(linked.cells)
    expect(opened.data.sheets[0]?.styles).toEqual(expect.arrayContaining(linked.styles))
    expect(opened.findings).toEqual([])
    // Other apps read the same links: ExcelJS, which Excel's own format rules shaped, finds every one.
    const other = new ExcelJS.Workbook()
    await (other.xlsx as unknown as { load(data: Uint8Array): Promise<unknown> }).load(saved)
    expect(['A1', 'B1', 'C1', 'E1'].map(address => (other.worksheets[0]?.getCell(address).value as ExcelJS.CellHyperlinkValue).hyperlink)).toEqual(targets.filter((_, column) => column !== 3))
  })

  it('reports threaded comments as dropped, keeping only the plain-note copy Excel saves with one', async () => {
    const source = new ExcelJS.Workbook()
    const sheet = source.addWorksheet('Sheet1')
    sheet.getCell('A1').value = 'copied'
    sheet.getCell('A1').note = '[Threaded comment] Comment: Ship by Friday'
    sheet.getCell('A2').value = 'threaded only'
    const entries = unzipSync(new Uint8Array(await source.xlsx.writeBuffer()))
    entries['xl/threadedComments/threadedComment1.xml'] = new TextEncoder().encode(
      '<ThreadedComments><threadedComment ref="A1"><text>Ship by Friday</text></threadedComment><threadedComment ref="A2"><text>Check the total</text></threadedComment></ThreadedComments>',
    )
    entries['xl/persons/person.xml'] = new TextEncoder().encode('<personList><person displayName="Sam"/></personList>')

    const { data, findings } = await readXlsx(zipSync(entries))
    const notes = data.sheets[0]?.styles.map(([column, row, style]) => [column, row, style.note])
    expect(notes).toEqual([[0, 0, '[Threaded comment] Comment: Ship by Friday']])
    expect(findings.find(finding => finding.construct === 'Threaded comments')).toMatchObject({
      severity: 'dropped',
      suggestedAlternative: expect.stringContaining('without that copy is not shown'),
    })
  })

  it('keeps the file number format, such as euros, through open and save', async () => {
    const source = new ExcelJS.Workbook()
    const sheet = source.addWorksheet('Costs')
    sheet.getCell('A1').value = 1234.5
    sheet.getCell('A1').numFmt = '[$€-2] #,##0.00'
    const opened = await readXlsx(new Uint8Array(await source.xlsx.writeBuffer()))
    const style = opened.data.sheets[0]?.styles[0]?.[2]
    expect(style).toMatchObject({ numberFormat: 'currency', formatCode: '[$€-2] #,##0.00' })
    expect(formatCellValue(1234.5, style ?? defaultCellStyle)).toContain('€')

    const styles = unzipSync(await writeXlsx(opened.data))['xl/styles.xml']
    expect(new TextDecoder().decode(styles)).toContain('formatCode="[$€-2] #,##0.00"')
  })

  it('shows custom format codes, and keeps the ones it cannot show unchanged', async () => {
    const codes = ['0.00E+00', '#,##0.00;[Red]-#,##0.00', 'dddd', '# ?/?']
    const source = new ExcelJS.Workbook()
    const sheet = source.addWorksheet('Codes')
    codes.forEach((code, row) => {
      sheet.getCell(row + 1, 1).value = -1234.5
      sheet.getCell(row + 1, 1).numFmt = code
    })
    const opened = await readXlsx(new Uint8Array(await source.xlsx.writeBuffer()))
    const reopened = await readXlsx(await writeXlsx(opened.data))
    const styles = reopened.data.sheets[0]?.styles.map(([, , style]) => style) ?? []

    expect(styles.map(style => style.formatCode)).toEqual(codes)
    expect(formatCellValue(-1234.5, styles[0] ?? defaultCellStyle)).toBe('-1.23E+03')
    expect(formatCellValue(-1234.5, styles[1] ?? defaultCellStyle)).toBe('-1,234.50')
    // Only the fraction is outside the supported codes.
    expect(opened.findings.map(finding => finding.construct)).toEqual(['Unmapped number formats'])
  })
})
