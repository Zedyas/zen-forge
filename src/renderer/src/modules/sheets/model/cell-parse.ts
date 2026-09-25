import type { CellInput, CellStyle } from './workbook-data'

function decimalPlaces(value: string): number {
  const decimal = value.match(/\.(\d+)/)?.[1]
  return decimal?.length ?? 0
}

/** Infers a typed value and number format from display text, as Excel and Google Sheets copy it. */
export function parseCellText(value: string): { readonly input: CellInput; readonly style?: CellStyle } {
  const trimmed = value.trim()
  if (trimmed === '') return { input: '' }
  if (trimmed.startsWith('=')) return { input: value }
  if (/^(true|false)$/i.test(trimmed)) return { input: trimmed.toLowerCase() === 'true' }

  const currency = trimmed.match(/^\(?\s*[$€£]\s*([\d,]+(?:\.\d+)?)\s*\)?$/)
  if (currency !== null) {
    const amount = Number(currency[1]?.replaceAll(',', '')) * (trimmed.startsWith('(') ? -1 : 1)
    return {
      input: amount,
      style: { numberFormat: 'currency', decimalPlaces: decimalPlaces(currency[1] ?? '') },
    }
  }

  const percent = trimmed.match(/^([+-]?[\d,]+(?:\.\d+)?)%$/)
  if (percent !== null) {
    return {
      input: Number(percent[1]?.replaceAll(',', '')) / 100,
      style: { numberFormat: 'percent', decimalPlaces: decimalPlaces(percent[1] ?? '') },
    }
  }

  if (/^[+-]?[\d,]+(?:\.\d+)?$/.test(trimmed)) {
    return {
      input: Number(trimmed.replaceAll(',', '')),
      style: trimmed.includes(',') || trimmed.includes('.')
        ? { numberFormat: 'number', decimalPlaces: decimalPlaces(trimmed) }
        : undefined,
    }
  }

  if (/^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$/.test(trimmed) && !Number.isNaN(Date.parse(trimmed))) {
    return { input: trimmed, style: { numberFormat: 'date', decimalPlaces: 0 } }
  }

  return { input: value }
}
