/**
 * CSV import and export.
 *
 * CSV carries no formatting, so a well-formed file imports losslessly; the only findings are for
 * guesses the importer had to make about the bytes themselves.
 */

import Papa from 'papaparse'
import type { ImportFindingInput } from '@shared/fidelity'
import { parseCellText } from '../model/cell-parse'
import { defaultColumnCount, defaultRowCount, type CellInput, type CellStyle, type SheetData, type WorkbookData } from '../model/workbook-data'

export interface CsvImport {
  readonly data: WorkbookData
  readonly findings: ImportFindingInput[]
}

interface DecodedText {
  readonly text: string
  /** Set when the bytes were not valid UTF-8 and a legacy encoding had to be assumed. */
  readonly assumedEncoding?: string
}

const delimiters = [',', ';', '\t', '|']

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

export function readCsv(bytes: Uint8Array, sheetName: string): CsvImport {
  const { text, assumedEncoding } = decode(bytes)
  // A file's final line break would otherwise count as a one-field row and skew delimiter detection.
  const parsed = Papa.parse<string[]>(text.replace(/\r?\n$/, ''), {
    header: false,
    delimitersToGuess: delimiters,
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
      construct: 'Malformed CSV rows',
      severity: 'degraded',
      location: `${parsed.errors.length} rows`,
    })
  }

  return { data: { sheets: [sheet], namedRanges: [] }, findings }
}

function quote(field: string): string {
  return /["\r\n,]/.test(field) ? `"${field.replaceAll('"', '""')}"` : field
}

/** RFC 4180 with a UTF-8 BOM, without which Excel for macOS mis-reads accented characters. */
export function writeCsv(rows: readonly (readonly string[])[]): Uint8Array {
  const body = rows.map(row => row.map(quote).join(',')).join('\r\n')
  return new TextEncoder().encode(`\uFEFF${body}`)
}
