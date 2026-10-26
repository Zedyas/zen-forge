import { describe, expect, it } from 'vitest'
import { shiftFormula } from './references'

describe('shiftFormula', () => {
  it.each([
    ['=A1+B1', 0, 1, '=A2+B2'],
    ['=$A$1+A$1+$A1', 2, 3, '=$A$1+C$1+$A4'],
    ['=SUM(A1:B2)*Sheet2!C3', 1, 0, '=SUM(B1:C2)*Sheet2!D3'],
    ["='Q1 ''26'!A1&\"A1\"", 0, 1, "='Q1 ''26'!A2&\"A1\""],
    ['=SUM(A:A)+SUM(1:2)', 1, 1, '=SUM(B:B)+SUM(2:3)'],
    ['=LOG10(A1)+1E5+Total', 0, 1, '=LOG10(A2)+1E5+Total'],
    ['=A1', 0, -1, '=#REF!'],
    ['plain A1 text', 1, 1, 'plain A1 text'],
  ])('%s moved %i across and %i down is %s', (formula, columns, rows, expected) => {
    expect(shiftFormula(formula, columns, rows)).toBe(expected)
  })
})
