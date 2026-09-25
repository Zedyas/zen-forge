import * as ExcelJS from 'exceljs'
import { unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { formatCellValue } from '../model/format'
import { defaultCellStyle, type WorkbookData } from '../model/workbook-data'
import { readXlsx, writeXlsx } from './xlsx'

/** A workbook that exercises every construct the format layer claims to preserve. */
const workbook: WorkbookData = {
  sheets: [
    {
      name: 'Sales',
      cells: [
        ['Region', 'Revenue', 'Share', 'Booked'],
        ['North', 42_100, '=B2/Totals!B1', 45_000],
        ['South', 38_750, '=B3/Totals!B1', 45_001],
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
        [0, 1, {
          numberFormat: 'general',
          decimalPlaces: 2,
          italic: true,
          underline: true,
          textColor: '#b91c1c',
        }],
        [1, 1, { numberFormat: 'currency', decimalPlaces: 2 }],
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
        hiddenRows: [2],
        hiddenColumns: [],
        freezeRows: 1,
        freezeColumns: 1,
      },
    },
    {
      name: 'Totals',
      cells: [['Target', 45_000]],
      values: [['Target', 45_000]],
      styles: [[1, 0, { numberFormat: 'number', decimalPlaces: 0 }]],
      layout: {
        rowCount: 1_000,
        columnCount: 26,
        columnWidths: [],
        hiddenRows: [],
        hiddenColumns: [],
        freezeRows: 0,
        freezeColumns: 0,
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

  it('reports macros, pivot tables, charts and merges as aggregated findings', async () => {
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

    const { findings } = await readXlsx(zipSync(entries))
    const byConstruct = new Map(findings.map(finding => [finding.construct, finding]))

    expect(byConstruct.get('Macros (VBA)')?.severity).toBe('dropped')
    expect(byConstruct.get('Pivot tables')?.severity).toBe('dropped')
    expect(byConstruct.get('Charts')?.severity).toBe('dropped')
    expect(byConstruct.get('Merged cells')).toEqual({
      construct: 'Merged cells',
      severity: 'degraded',
      location: 'Sheet1 (2 ranges)',
    })
    expect(findings.length).toBe(byConstruct.size)
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
})
