import { unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { readXlsx, writeXlsx } from './io'
import { textCellInput } from './model/cell-parse'
import { Workbook } from './model/Workbook'

describe('workbook .xlsx round trip', () => {
  it('keeps formulas, computed values, styles, layout and names through export and re-import', async () => {
    const original = new Workbook()
    const sheetId = original.sheets()[0]?.id ?? 0
    original.setBlock(sheetId, [0, 0], [['Item', 'Cost'], ['Rent', '$2,150.00'], ['Food', '$612.50']])
    original.setCellInput(sheetId, [1, 3], '=SUM(B2:B3)')
    original.updateStyle(sheetId, [{ x: 0, y: 0, width: 2, height: 1 }], { bold: true, fillColor: '#dcf2e6' })
    original.updateStyle(sheetId, [{ x: 1, y: 3, width: 1, height: 1 }], { textColor: '#c9352b', align: 'center', underline: true })
    original.setNumberFormat(sheetId, [{ x: 1, y: 3, width: 1, height: 1 }], 'currency')
    original.freezeRows(sheetId, 1)
    original.setColumnWidth(sheetId, 0, 180)
    original.addNamedRange('Costs', sheetId, { x: 1, y: 1, width: 1, height: 2 })

    const { data, findings } = await readXlsx(await writeXlsx(original.toData()))
    const reopened = new Workbook(data)
    const reopenedId = reopened.sheets()[0]?.id ?? 0

    expect(findings).toEqual([])
    expect(reopened.getCellInput(reopenedId, [1, 3])).toBe('=SUM(B2:B3)')
    expect(reopened.getCellValue(reopenedId, [1, 3])).toBe(2762.5)
    expect(reopened.getStyle(reopenedId, [0, 0])).toMatchObject({ bold: true, fillColor: '#dcf2e6' })
    expect(reopened.getStyle(reopenedId, [1, 3])).toMatchObject({
      numberFormat: 'currency', decimalPlaces: 2, textColor: '#c9352b', align: 'center', underline: true,
    })
    expect(reopened.getFrozenRowCount(reopenedId)).toBe(1)
    expect(reopened.getColumnWidth(reopenedId, 0)).toBe(180)
    expect(reopened.namedRanges().map(named => named.name)).toEqual(['Costs'])
    expect(reopened.isDirty()).toBe(false)
  })

  it('keeps text cells that look like numbers, booleans or formulas as text', async () => {
    const original = new Workbook()
    const sheetId = original.sheets()[0]?.id ?? 0
    const texts = ['00123', '=x', 'TRUE', '1234567890123456789']
    texts.forEach((text, col) => original.setCellInput(sheetId, [col, 0], textCellInput(text)))

    const written = await writeXlsx(original.toData())
    const strings = new TextDecoder().decode(unzipSync(written)['xl/sharedStrings.xml'])
    const reopened = new Workbook((await readXlsx(written)).data)
    const reopenedId = reopened.sheets()[0]?.id ?? 0

    // The file stores the text itself; the apostrophe is only the engine's marker.
    expect(texts.every(text => strings.includes(`<t>${text}</t>`))).toBe(true)
    expect(texts.map((_, col) => reopened.getCellValue(reopenedId, [col, 0]))).toEqual(texts)
  })
})
