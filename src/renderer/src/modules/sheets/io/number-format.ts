/** Translation between Excel number format patterns and the workbook's five formats. */

import { defaultCellStyle, type CellStyle, type NumberFormat } from '../model/workbook-data'

export interface ParsedNumberFormat {
  readonly numberFormat: NumberFormat
  readonly decimalPlaces: number
  /** False when the pattern carries detail the model drops: extra sections, text or scientific formats. */
  readonly mapped: boolean
}

const general: ParsedNumberFormat = { numberFormat: 'general', decimalPlaces: 2, mapped: true }
const unmapped: ParsedNumberFormat = { numberFormat: 'general', decimalPlaces: 2, mapped: false }

/** `[$-409]` names a locale; `[$€-2]` and `[$$-409]` name a currency. */
const localeMarker = /\[\$-[0-9A-Fa-f]+\]/g
const bracketed = /\[[^\]]*\]/g
const quoted = /"[^"]*"/g
const escaped = /\\./g

/** Splits `positive;negative;zero;text` while ignoring separators inside quotes or brackets. */
function splitSections(pattern: string): readonly string[] {
  const sections: string[] = []
  let current = ''
  let quoting = false
  let depth = 0

  for (const character of pattern) {
    if (character === '"') quoting = !quoting
    if (!quoting && character === '[') depth += 1
    if (!quoting && character === ']') depth = Math.max(0, depth - 1)
    if (character === ';' && !quoting && depth === 0) {
      sections.push(current)
      current = ''
      continue
    }
    current += character
  }

  sections.push(current)
  return sections
}

export function parseNumberFormat(pattern: string): ParsedNumberFormat {
  const trimmed = pattern.trim()
  if (trimmed === '' || trimmed.toLowerCase() === 'general') return general

  const sections = splitSections(trimmed)
  const positive = sections[0] ?? trimmed
  const withoutLocale = positive.replace(localeMarker, '')
  const literal = withoutLocale.replace(bracketed, '').replace(quoted, '').replace(escaped, '')
  const decimalPlaces = literal.match(/\.(0+)/)?.[1].length ?? 0
  const mapped = sections.length === 1

  // Text placeholders and scientific notation have no equivalent; the value still reads as a number.
  if (literal.includes('@') || /e[+-]?[0#]/i.test(literal)) return unmapped

  if (literal.includes('%')) return { numberFormat: 'percent', decimalPlaces, mapped }
  if (withoutLocale.includes('[$') || /[$€£¥]/.test(withoutLocale)) {
    return { numberFormat: 'currency', decimalPlaces, mapped }
  }
  if (/[dy]/i.test(literal)) return { numberFormat: 'date', decimalPlaces: 0, mapped }
  if (/[0#]/.test(literal) && (literal.includes(',') || decimalPlaces > 0)) {
    return { numberFormat: 'number', decimalPlaces, mapped }
  }

  return unmapped
}

/**
 * A file's number format pattern as a cell style: the nearest of the five formats, plus the pattern
 * itself when it says more than the one this app would write for that style. `mapped` is false when
 * the model can only approximate the pattern.
 */
export function numberFormatStyle(pattern: string): { readonly style: CellStyle; readonly mapped: boolean } {
  const parsed = parseNumberFormat(pattern)
  const style: CellStyle = { ...defaultCellStyle, numberFormat: parsed.numberFormat, decimalPlaces: parsed.decimalPlaces }
  const keepsPattern = pattern !== numberFormatPattern(style) && !(parsed.mapped && parsed.numberFormat === 'general')
  return { style: keepsPattern ? { ...style, formatCode: pattern } : style, mapped: parsed.mapped }
}

/** The pattern written for a style, or undefined when Excel's General format already matches. */
export function numberFormatPattern(style: CellStyle): string | undefined {
  const places = Math.max(0, Math.min(12, Math.round(style.decimalPlaces)))
  const decimals = places > 0 ? `.${'0'.repeat(places)}` : ''

  switch (style.numberFormat) {
    case 'currency':
      return `$#,##0${decimals}`
    case 'percent':
      return `0${decimals}%`
    case 'date':
      return 'yyyy-mm-dd'
    case 'number':
      return `#,##0${decimals}`
    case 'general':
      return undefined
  }
}
