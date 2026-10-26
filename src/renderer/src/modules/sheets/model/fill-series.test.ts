import { describe, expect, it } from 'vitest'
import { fillSeries, type FillSeed } from './fill-series'

function seeds(...values: readonly (string | number)[]): FillSeed[] {
  return values.map(value => ({ input: value, value, date: false }))
}

function inputs(from: readonly FillSeed[], positions: readonly number[]): readonly unknown[] {
  return fillSeries(from, positions).map(cell => cell.input)
}

describe('fillSeries', () => {
  it.each([
    ['two numbers continue by their step', seeds(1, 2), [2, 3], [3, 4]],
    ['tens continue', seeds(10, 20), [2], [30]],
    ['decimals continue without float noise', seeds(0.1, 0.2), [2], [0.3]],
    ['one number repeats', seeds(7), [1, 2], [7, 7]],
    ['a fill up counts down', seeds(3, 4), [-1, -2], [2, 1]],
    ['short months keep their case', seeds('JAN'), [1, 2], ['FEB', 'MAR']],
    ['long months wrap', seeds('November'), [1, 2], ['December', 'January']],
    ['months step by the seed step', seeds('Jan', 'Mar'), [2], ['May']],
    ['weekdays continue', seeds('mon', 'tue'), [2], ['wed']],
    ['numbered text counts on, keeping zeros', seeds('Item 09'), [1], ["'Item 10"]],
    ['quarters wrap after Q4', seeds('Q3', 'Q4'), [2], ["'Q1"]],
    ['long digit runs count on exactly', seeds('12345678901234567890'), [1, -1], ["'12345678901234567891", "'12345678901234567889"]],
    ['other text repeats as a pattern', seeds('a', 'b'), [2, 3, -1], ['a', 'b', 'b']],
  ])('%s', (_, from, positions, expected) => {
    expect(inputs(from, positions)).toEqual(expected)
  })

  it('steps a date by a day, and dates a month apart by months', () => {
    const date = (serial: number): FillSeed => ({ input: serial, value: serial, date: true })
    // 2026-01-31 and 2026-03-31 are two months apart, so May 31 comes next.
    expect(inputs([date(46_053)], [1])).toEqual([46_054])
    expect(inputs([date(46_053), date(46_112)], [2])).toEqual([46_173])
  })

  it('copies formulas instead of continuing them, naming the seed to copy from', () => {
    const from: FillSeed[] = [{ input: '=A1*2', value: 2, date: false }, { input: 3, value: 3, date: false }]
    expect(fillSeries(from, [2, 3])).toEqual([{ seed: 0, input: '=A1*2' }, { seed: 1, input: 3 }])
  })
})
