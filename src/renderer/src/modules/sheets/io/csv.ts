/**
 * Delimited text import and export: CSV (comma-separated) and TSV (tab-separated).
 *
 * Neither carries formatting, so a well-formed file imports losslessly; the only findings are for
 * guesses the importer had to make about the bytes themselves.
 */

import Papa from 'papaparse'
import type { ImportFindingInput } from '@shared/fidelity'
import { parseCellText } from '../model/cell-parse'
import { defaultColumnCount, defaultRowCount, type CellInput, type CellStyle, type SheetData, type WorkbookData } from '../model/workbook-data'

export type DelimitedFormat = 'csv' | 'tsv'

export interface DelimitedImport {
  readonly data: WorkbookData
  readonly findings: ImportFindingInput[]
}

export function isDelimitedFormat(extension: string): extension is DelimitedFormat {
  return extension === 'csv' || extension === 'tsv'
}

interface DecodedText {
  readonly text: string
  /** Set when the bytes were not valid UTF-8 and a legacy encoding had to be assumed. */
  readonly assumedEncoding?: string
}

/** Separators a CSV may use: many locales write `;` because `,` is their decimal mark. */
const csvDelimiters = [',', ';', '\t', '|']
const separators: Record<DelimitedFormat, string> = { csv: ',', tsv: '\t' }

function startsWith(bytes: Uint8Array, ...prefix: readonly number[]): boolean {
  return prefix.every((byte, index) => bytes[index] === byte)
}

function decode(bytes: Uint8Array): DecodedText {
  if (startsWith(bytes, 0xef, 0xbb, 0xbf)) {
    return { text: new TextDecoder('utf-8').decode(bytes.subarray(3)) }
  }
  if (startsWith(bytes, 0xff, 0xfe)) {
    return { text: new TextDecoder('utf-16le').decode(bytes.subarray(2)) }
  }
  if (startsWith(bytes, 0xfe, 0xff)) {
    return { text: new TextDecoder('utf-16be').decode(bytes.subarray(2)) }
  }

  try {
    return { text: new TextDecoder('utf-8', { fatal: true }).decode(bytes) }
  } catch {
    return { text: new TextDecoder('windows-1252').decode(bytes), assumedEncoding: 'windows-1252' }
  }
}

/** Trailing blank lines are an artefact of the file's final newline, not data. */
function withoutTrailingBlanks(rows: readonly (readonly string[])[]): readonly (readonly string[])[] {
  let end = rows.length
  while (end > 0 && (rows[end - 1] ?? []).every(field => field === '')) end -= 1
  return rows.slice(0, end)
}

/** A CSV's separator is guessed; a TSV's is always a tab, so a comma inside a field stays text. */
export function readDelimited(bytes: Uint8Array, sheetName: string, format: DelimitedFormat): DelimitedImport {
  const { text, assumedEncoding } = decode(bytes)
  // A file's final line break would otherwise count as a one-field row and skew delimiter detection.
  const parsed = Papa.parse<string[]>(text.replace(/\r?\n$/, ''), {
    header: false,
    ...(format === 'tsv' ? { delimiter: separators.tsv } : { delimitersToGuess: csvDelimiters }),
    skipEmptyLines: false,
  })
  const rows = withoutTrailingBlanks(parsed.data)

  const cells: CellInput[][] = []
  const styles: [number, number, CellStyle][] = []
  let columnCount = 0

  rows.forEach((row, rowIndex) => {
    columnCount = Math.max(columnCount, row.length)
    cells.push(row.map((field, columnIndex) => {
      const { input, style } = parseCellText(field)
      if (style !== undefined) styles.push([columnIndex, rowIndex, style])
      return input
    }))
  })

  const sheet: SheetData = {
    name: sheetName,
    cells,
    styles,
    layout: {
      rowCount: Math.max(defaultRowCount, rows.length),
      columnCount: Math.max(defaultColumnCount, columnCount),
      columnWidths: [],
      hiddenRows: [],
      hiddenColumns: [],
      freezeRows: 0,
      freezeColumns: 0,
    },
  }

  const findings: ImportFindingInput[] = []
  if (assumedEncoding !== undefined) {
    findings.push({
      construct: 'Non-Unicode text encoding',
      severity: 'degraded',
      location: `Read as ${assumedEncoding}`,
      suggestedAlternative: 'Re-export the file as UTF-8 if any characters look wrong.',
    })
  }
  if (parsed.errors.length > 0) {
    findings.push({
      construct: `Malformed ${format.toUpperCase()} rows`,
      severity: 'degraded',
      location: `${parsed.errors.length} rows`,
    })
  }

  return { data: { sheets: [sheet], namedRanges: [] }, findings }
}

/**
 * RFC 4180 with a UTF-8 BOM, without which Excel for macOS mis-reads accented characters. A field is
 * quoted only when it holds the separator, a quote or a line break.
 */
export function writeDelimited(rows: readonly (readonly string[])[], format: DelimitedFormat): Uint8Array {
  const separator = separators[format]
  const quote = (field: string): string =>
    field.includes(separator) || /["\r\n]/.test(field) ? `"${field.replaceAll('"', '""')}"` : field
  const body = rows.map(row => row.map(quote).join(separator)).join('\r\n')
  return new TextEncoder().encode(`\uFEFF${body}`)
}
