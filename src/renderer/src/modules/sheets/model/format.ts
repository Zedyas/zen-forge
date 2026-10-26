import type { CellStyle } from './workbook-data'

const dayMilliseconds = 86_400_000
/** Excel's day zero (serial 0 = 1899-12-30), which keeps the 1900 leap-year quirk compatible. */
const serialEpoch = Date.UTC(1899, 11, 30)

/** The ISO code for display, from the file's currency symbol; `$` and anything unrecognised show as dollars. */
function currencyCode(formatCode: string | undefined): string {
  if (formatCode?.includes('€')) return 'EUR'
  if (formatCode?.includes('£')) return 'GBP'
  if (formatCode?.includes('¥')) return 'JPY'
  return 'USD'
}

/** Whether the style shows numbers as dates, so they fill and export as dates. */
export function isDateStyle(style: CellStyle): boolean {
  return style.numberFormat === 'date'
}

/** Formats a computed value the way the grid shows it and a CSV export writes it. */
export function formatCellValue(value: string | number | boolean | null, style: CellStyle): string {
  if (value === null) return ''
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  const digits = { minimumFractionDigits: style.decimalPlaces, maximumFractionDigits: style.decimalPlaces }
  if (typeof value === 'number') {
    switch (style.numberFormat) {
      case 'currency':
        return value.toLocaleString(undefined, { style: 'currency', currency: currencyCode(style.formatCode), ...digits })
      case 'percent':
        return value.toLocaleString(undefined, { style: 'percent', ...digits })
      case 'number':
        return value.toLocaleString(undefined, digits)
      case 'date': {
        const date = new Date(serialEpoch + value * dayMilliseconds)
        return Number.isNaN(date.valueOf()) ? String(value) : date.toLocaleDateString(undefined, { timeZone: 'UTC' })
      }
      case 'general':
        // Like Excel's General: no thousands separators, so copied or exported text reads back as the same number.
        return value.toLocaleString(undefined, { maximumFractionDigits: 8, useGrouping: false })
    }
  }
  if (style.numberFormat === 'date' && typeof value === 'string') {
    // `new Date('2024-03-05')` is midnight UTC, the day before west of Greenwich; a date alone means that day here.
    const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value)
    const date = iso === null ? new Date(value) : new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]))
    if (!Number.isNaN(date.valueOf())) return date.toLocaleDateString()
  }
  return value
}

/** A date serial as ISO 8601 (`2026-08-13`), with the time (`2026-08-13 14:30:00`) when it has one. */
function isoDate(serial: number): string {
  const date = new Date(serialEpoch + Math.round(serial * dayMilliseconds / 1000) * 1000)
  if (Number.isNaN(date.valueOf())) return String(serial)
  const [day, time] = date.toISOString().replace(/\.\d+Z$/, '').split('T')
  return time === '00:00:00' ? `${day}` : `${day} ${time}`
}

/**
 * A computed value as CSV and TSV store it: the number at full precision, not rounded to the shown
 * decimals, and dates as ISO 8601, so the file reads back as the same values in any app or locale.
 */
export function delimitedValue(value: string | number | boolean | null, style: CellStyle): string {
  if (value === null) return ''
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (typeof value === 'number' && isDateStyle(style)) return isoDate(value)
  return String(value)
}

/**
 * A selection statistic (sum, average…) in the active cell's format when that format is numeric, so a
 * currency column totals in currency. General and date cells fall back to two decimal places.
 */
export function formatStatistic(value: number, style: CellStyle): string {
  if (style.numberFormat === 'general' || style.numberFormat === 'date') {
    return value.toLocaleString(undefined, { maximumFractionDigits: 2 })
  }
  return formatCellValue(value, style)
}
