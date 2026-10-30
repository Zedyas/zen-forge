import { textCellInput } from './cell-parse'
import type { CellInput } from './workbook-data'

/** A source cell of a fill, as the series rules read it. */
export interface FillSeed {
  /** The raw input; a formula is copied (with its references moved), never continued. */
  readonly input: CellInput
  readonly value: string | number | boolean | null
  /** The cell shows a date, so its number steps by days or months. */
  readonly date: boolean
}

export interface FilledCell {
  /** The seed whose style the cell takes, and whose formula it copies when the fill repeats. */
  readonly seed: number
  readonly input: CellInput
}

const dayMilliseconds = 86_400_000
const serialEpoch = Date.UTC(1899, 11, 30)

/** Name lists that fill as cycles. Short names come first, so a lone `May` continues as `Jun`. */
const cycles: readonly (readonly string[])[] = [
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
  ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
]

function modulo(value: number, divisor: number): number {
  return ((value % divisor) + divisor) % divisor
}

/** Least-squares line through (0, values[0]), (1, values[1])…, as Excel's linear trend fills. */
function trend(values: readonly number[]): { readonly start: number; readonly step: number } {
  const meanX = (values.length - 1) / 2
  const meanY = values.reduce((total, value) => total + value, 0) / values.length
  let numerator = 0
  let denominator = 0
  values.forEach((value, x) => {
    numerator += (x - meanX) * (value - meanY)
    denominator += (x - meanX) ** 2
  })
  const step = denominator === 0 ? 0 : numerator / denominator
  return { start: meanY - step * meanX, step }
}

/** Rounds away float noise (0.1 + 0.2), which a fill would otherwise write into the cell. */
function clean(value: number): number {
  return Number(value.toPrecision(15))
}

/** The one step between every pair of neighbours, or undefined when the steps differ. */
function evenStep(values: readonly number[], cycle?: number): number | undefined {
  if (values.length === 1) return 1
  const steps = values.slice(1).map((value, index) => value - (values[index] ?? 0)).map(step => cycle === undefined ? step : modulo(step, cycle))
  return steps.every(step => step === steps[0]) && steps[0] !== 0 ? steps[0] : undefined
}

function sameCase(name: string, model: string): string {
  if (model === model.toUpperCase()) return name.toUpperCase()
  if (model === model.toLowerCase()) return name.toLowerCase()
  return name
}

function monthParts(serial: number): { readonly year: number; readonly month: number; readonly day: number } {
  const date = new Date(serialEpoch + serial * dayMilliseconds)
  return { year: date.getUTCFullYear(), month: date.getUTCMonth(), day: date.getUTCDate() }
}

/** Dates on the same day of the month an even number of months apart continue by months, as in Excel. */
function monthSeries(serials: readonly number[]): ((position: number) => number) | undefined {
  if (serials.length < 2 || serials.some(serial => !Number.isInteger(serial))) return undefined
  const parts = serials.map(monthParts)
  const first = parts[0]
  if (first === undefined || parts.some(part => part.day !== first.day)) return undefined
  const step = evenStep(parts.map(part => part.year * 12 + part.month))
  if (step === undefined) return undefined
  return position => {
    const month = first.month + step * position
    // Day 31 in a 30-day month lands on its last day.
    const lastDay = new Date(Date.UTC(first.year, month + 1, 0)).getUTCDate()
    return (Date.UTC(first.year, month, Math.min(first.day, lastDay)) - serialEpoch) / dayMilliseconds
  }
}

function numberSeries(seeds: readonly FillSeed[]): ((position: number) => CellInput) | undefined {
  const values = seeds.map(seed => seed.value)
  if (!values.every(value => typeof value === 'number')) return undefined
  const numbers = values.filter(value => typeof value === 'number')
  if (seeds.every(seed => seed.date)) {
    const months = monthSeries(numbers)
    if (months !== undefined) return months
    // One date steps by a day.
    if (numbers.length === 1) return position => (numbers[0] ?? 0) + position
  }
  // One number repeats, as in Excel without the Option key.
  if (numbers.length === 1) return undefined
  const { start, step } = trend(numbers)
  return position => clean(start + step * position)
}

function cycleSeries(texts: readonly string[]): ((position: number) => CellInput) | undefined {
  for (const names of cycles) {
    const lower = names.map(name => name.toLowerCase())
    const indexes = texts.map(text => lower.indexOf(text.toLowerCase()))
    if (indexes.some(index => index < 0)) continue
    const step = evenStep(indexes, names.length)
    const first = indexes[0] ?? 0
    const model = texts[0] ?? ''
    if (step === undefined) return undefined
    return position => sameCase(names[modulo(first + step * position, names.length)] ?? '', model)
  }
  return undefined
}

/** Text ending in a number (Item 1, Week 09, Q1) counts on; quarters wrap after Q4. */
function numberedTextSeries(texts: readonly string[]): ((position: number) => CellInput) | undefined {
  const parts = texts.map(text => text.match(/^(.*?)(\d+)$/))
  const prefix = parts[0]?.[1]
  if (prefix === undefined || parts.some(part => part === null || part[1] !== prefix)) return undefined
  const digits = parts.map(part => part?.[2] ?? '')
  const width = digits[0]?.startsWith('0') === true ? (digits[0]?.length ?? 0) : 0
  const text = (number: number | bigint): CellInput => textCellInput(`${prefix}${String(number).padStart(width, '0')}`)
  const quarters = digits.map(Number)
  if (/^(?:q|qtr|quarter)\s*$/i.test(prefix) && quarters.every(number => number >= 1 && number <= 4)) {
    const step = evenStep(quarters.map(number => number - 1), 4)
    const first = quarters[0] ?? 1
    return step === undefined ? undefined : position => text(modulo(first - 1 + step * position, 4) + 1)
  }
  // A run of digits can be longer than a float holds exactly (a 20-digit ID), so it counts as a BigInt.
  const numbers = digits.map(digit => BigInt(digit))
  const steps = numbers.slice(1).map((number, index) => number - (numbers[index] ?? 0n))
  const step = steps.length === 0 ? 1n : steps.every(value => value === steps[0]) && steps[0] !== 0n ? steps[0] : undefined
  if (step === undefined) return undefined
  const first = numbers[0] ?? 0n
  return position => {
    const number = first + step * BigInt(position)
    return text(number < 0n ? -number : number)
  }
}

function textSeries(seeds: readonly FillSeed[]): ((position: number) => CellInput) | undefined {
  const values = seeds.map(seed => seed.value)
  if (!values.every(value => typeof value === 'string' && value !== '')) return undefined
  const texts = values.filter(value => typeof value === 'string')
  return cycleSeries(texts) ?? numberedTextSeries(texts)
}

/**
 * What the fill handle writes at each position. Seeds sit at positions 0 to n - 1; a fill down or right
 * continues at n, n + 1…, and a fill up or left at -1, -2…. Numbers (from two seeds), dates, month and
 * weekday names and numbered text continue as a series; anything else, formulas included, repeats.
 */
export function fillSeries(seeds: readonly FillSeed[], positions: readonly number[]): readonly FilledCell[] {
  const continuing = seeds.length === 0 || seeds.some(seed => typeof seed.input === 'string' && seed.input.startsWith('='))
    ? undefined
    : numberSeries(seeds) ?? textSeries(seeds)
  return positions.map(position => {
    const seed = modulo(position, Math.max(1, seeds.length))
    return { seed, input: continuing === undefined ? seeds[seed]?.input ?? null : continuing(position) }
  })
}
