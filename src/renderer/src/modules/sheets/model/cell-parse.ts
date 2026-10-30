import { dayMilliseconds, serialEpoch } from './format'
import type { CellInput, CellStyle } from './workbook-data'

function decimalPlaces(value: string): number {
  const decimal = value.match(/\.(\d+)/)?.[1]
  return decimal?.length ?? 0
}

/**
 * Text that typing would turn into something else: a number, date, time, percentage, boolean, error or
 * formula. Deliberately broad; a false positive only costs an apostrophe in the formula bar.
 */
const readsAsValue = /^\s*(?:[-+]?[$€£¥]?\s*[-+]?[\d.]+(?:e[-+]?\d+)?\s*[%$€£¥]?|[\d\s/:.-]+(?:\s*[ap]m)?|true|false|#(?:null!|div\/0!|value!|ref!|name\?|num!|n\/a|cycle!|error!|spill!|lic!))\s*$/i

/**
 * The raw input that keeps `text` as text, whatever the engine's parser would make of it. An apostrophe
 * marks it, as in Excel and Google Sheets; the cell's value is the text without it.
 */
export function textCellInput(text: string): string {
  return `'${text}`
}

/**
 * A raw input as the formula bar shows it and Find searches it: a text cell keeps its apostrophe only
 * where typing the text alone would read as something else, so plain words show as they are.
 */
export function shownInput(input: string): string {
  const text = input.slice(1)
  return input.startsWith("'") && !text.startsWith('=') && !text.startsWith("'") && !readsAsValue.test(text) ? text : input
}

/** Digits, with any commas grouping by thousands (1,234 or 12,345,678); 1,5 is not a number. */
function groupedNumber(digits: string): number | undefined {
  if (!/^(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$/.test(digits)) return undefined
  return Number(digits.replaceAll(',', ''))
}

/**
 * A plain number, unless reading it as one would change it: a leading zero (00123, a ZIP code or ID)
 * would be dropped, and more than 15 integer digits (card numbers, long IDs) cannot be stored exactly.
 */
function plainNumber(text: string): number | undefined {
  // An exponent needs its sign (1e+21, as numbers print); a code like 12E3 stays text.
  const match = text.match(/^[+-]?(\d*)(?:\.\d*)?(?:e[-+]\d+)?$/i)
  const integer = match?.[1]
  if (integer === undefined || !/\d/.test(text.replace(/e.*$/i, ''))) return undefined
  if (/^0\d/.test(integer) || integer.replace(/^0+/, '').length > 15) return undefined
  return Number(text)
}

/**
 * An ISO date (`2026-08-13`), or one with a time (`2026-08-13 14:30:00`, or with a `T` and without
 * seconds), as a CSV export writes one, as a date serial of that wall-clock time; undefined for anything
 * else. The engine reads neither form itself.
 */
function isoDateTimeSerial(text: string): number | undefined {
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/)
  if (match === null) return undefined
  const [year = 0, month = 0, day = 0, hour = 0, minute = 0, second = 0] = match.slice(1).map(part => Number(part ?? 0))
  const time = Date.UTC(year, month - 1, day, hour, minute, second)
  const date = new Date(time)
  if (date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day || hour > 23 || minute > 59 || second > 59) return undefined
  return (time - serialEpoch) / dayMilliseconds
}

/**
 * Infers a typed value and number format from display text, as Excel and Google Sheets copy it, and as
 * a CSV holds it. Text that would lose information as a number (00123, 1,5, a 16-digit ID) stays text.
 */
export function parseCellText(value: string): { readonly input: CellInput; readonly style?: CellStyle } {
  const trimmed = value.trim()
  // An empty field is an empty cell, as when Excel pastes or opens one.
  if (trimmed === '') return { input: null }
  if (trimmed.startsWith('=')) return { input: value }
  if (/^(true|false)$/i.test(trimmed)) return { input: trimmed.toLowerCase() === 'true' }

  // A negative amount shows as -$12.50 (Excel, Google Sheets) or ($12.50) (accounting).
  const currency = trimmed.match(/^-?\(?\s*[$€£]\s*([\d,]+(?:\.\d+)?)\s*\)?$/)
  const currencyAmount = currency === null ? undefined : groupedNumber(currency[1] ?? '')
  if (currency !== null && currencyAmount !== undefined) {
    return {
      input: currencyAmount * (trimmed.startsWith('(') || trimmed.startsWith('-') ? -1 : 1),
      style: { numberFormat: 'currency', decimalPlaces: decimalPlaces(currency[1] ?? '') },
    }
  }

  const percent = trimmed.match(/^([+-]?)([\d,]+(?:\.\d+)?)%$/)
  const percentAmount = percent === null ? undefined : groupedNumber(percent[2] ?? '')
  if (percent !== null && percentAmount !== undefined) {
    return {
      input: (percent[1] === '-' ? -percentAmount : percentAmount) / 100,
      style: { numberFormat: 'percent', decimalPlaces: decimalPlaces(percent[2] ?? '') },
    }
  }

  const grouped = trimmed.match(/^([+-]?)(\d{1,3}(?:,\d{3})+(?:\.\d+)?)$/)
  const groupedAmount = grouped === null ? undefined : groupedNumber(grouped[2] ?? '')
  if (grouped !== null && groupedAmount !== undefined) {
    return {
      input: grouped[1] === '-' ? -groupedAmount : groupedAmount,
      style: { numberFormat: 'number', decimalPlaces: decimalPlaces(trimmed) },
    }
  }

  const number = plainNumber(trimmed)
  if (number !== undefined) {
    return {
      input: number,
      style: trimmed.includes('.') && !/e/i.test(trimmed) ? { numberFormat: 'number', decimalPlaces: decimalPlaces(trimmed) } : undefined,
    }
  }

  const dateTime = isoDateTimeSerial(trimmed)
  if (dateTime !== undefined) return { input: dateTime, style: { numberFormat: 'date', decimalPlaces: 0 } }

  if (/^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$/.test(trimmed) && !Number.isNaN(Date.parse(trimmed))) {
    return { input: trimmed, style: { numberFormat: 'date', decimalPlaces: 0 } }
  }

  return { input: textCellInput(value) }
}
