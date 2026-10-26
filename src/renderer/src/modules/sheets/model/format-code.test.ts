import { describe, expect, it } from 'vitest'
import { formatWithCode, isDateCode, supportsFormatCode } from './format-code'

// 2026-03-07 14:05:09, a Saturday, as an Excel serial.
const dateTime = 46_088 + (14 * 3600 + 5 * 60 + 9) / 86_400

describe('formatWithCode', () => {
  it.each([
    ['0', 1234.5, '1235'],
    ['0.00', 0.5, '0.50'],
    ['#,##0', 1_234_567, '1,234,567'],
    ['#,##0.00', -1234.5, '-1,234.50'],
    ['#,###', 0, ''],
    ['000', 5, '005'],
    ['#.##', 3.1, '3.1'],
    ['0%', 0.256, '26%'],
    ['0.00%', 0.256, '25.60%'],
    ['$#,##0.00', 1234.5, '$1,234.50'],
    ['[$€-407] #,##0.00', 1234.5, '€ 1,234.50'],
    ['"Qty "0', 12, 'Qty 12'],
    ['0.0,,"M"', 2_500_000, '2.5M'],
    ['0.00E+00', 12_345, '1.23E+04'],
    ['0.00E+00', 0.000_123, '1.23E-04'],
    ['#,##0.00;(#,##0.00)', -5, '(5.00)'],
    ['0;-0;"zero"', 0, 'zero'],
    ['0;-0;0;"Text: "@', 'abc', 'Text: abc'],
    ['0.00', 'words', 'words'],
    ['yyyy-mm-dd', dateTime, '2026-03-07'],
    ['dd/mm/yy hh:mm:ss', dateTime, '07/03/26 14:05:09'],
    ['mmm d, yyyy', dateTime, 'Mar 7, 2026'],
    ['mmmm', dateTime, 'March'],
    ['ddd', dateTime, 'Sat'],
    ['dddd', dateTime, 'Saturday'],
    ['h:mm AM/PM', dateTime, '2:05 PM'],
    ['h:mm am/pm', 0.25, '6:00 am'],
    // Excel rounds on the value's 15 significant digits, half away from zero: 2.675 is 2.67499… in binary.
    ['0.00', 2.675, '2.68'],
    ['0.00', 1.005, '1.01'],
    ['0.0', -0.05, '-0.1'],
    ['[Red]General', -0.1, '-0.1'],
  ])('shows %s on %s as %s', (code, value, text) => {
    expect(formatWithCode(value, code)?.text).toBe(text)
  })

  it('colours the section a value falls in', () => {
    expect(formatWithCode(-3, '0.0;[Red]-0.0')).toEqual({ text: '-3.0', color: '#c9352b' })
    expect(formatWithCode(3, '0.0;[Red]-0.0')).toEqual({ text: '3.0' })
  })

  it('leaves codes outside the subset to General, and knows date codes', () => {
    // Excel allows 30 decimal places; more would make the formatter throw.
    expect(supportsFormatCode(`0.${'0'.repeat(101)}`)).toBe(false)
    expect(formatWithCode(1.5, `0.${'0'.repeat(101)}`)).toBeUndefined()
    expect(formatWithCode(1.5, `0.${'0'.repeat(30)}`)?.text).toBe(`1.5${'0'.repeat(29)}`)
    expect(supportsFormatCode('[>100]0.0')).toBe(false)
    expect(supportsFormatCode('# ?/?')).toBe(false)
    expect(supportsFormatCode('[h]:mm')).toBe(false)
    expect(formatWithCode(1, '[>100]0.0')).toBeUndefined()
    expect(isDateCode('h:mm AM/PM')).toBe(true)
    expect(isDateCode('#,##0.00')).toBe(false)
  })
})
