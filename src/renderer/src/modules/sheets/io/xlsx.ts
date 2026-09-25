/**
 * xlsx import and export.
 *
 * The importer is deliberately pessimistic: every construct in the file that the workbook model
 * cannot hold raises a finding, because a file reported as lossless is one the user will overwrite.
 */

import * as ExcelJS from 'exceljs'
import { HyperFormula } from 'hyperformula'
import type { ImportFindingInput } from '@shared/fidelity'
import { sheetRangeReference } from '../model/address'
import {
  defaultCellStyle,
  defaultColumnCount,
  defaultRowCount,
  isDefaultStyle,
  type CellInput,
  type CellStyle,
  type NamedRange,
  type SheetData,
  type WorkbookData,
} from '../model/workbook-data'
import { aggregateFindings, type FindingEvent } from './findings'
import { numberFormatPattern, numberFormatStyle } from './number-format'
import { assertZipFitsInMemory } from '../../../services/zip'
import { packageFindings, readWorkbookPackage, type DefaultFont } from './package-parts'

export interface XlsxImport {
  readonly data: WorkbookData
  readonly findings: ImportFindingInput[]
}

/** ExcelJS declares Node Buffers; the zip layer underneath works on any byte array. */
interface ByteXlsx {
  load(data: Uint8Array): Promise<unknown>
  writeBuffer(): Promise<Uint8Array>
}

/** Details ExcelJS parses into its models but leaves out of its published types. */
interface WorksheetDetail {
  readonly conditionalFormattings?: readonly unknown[]
  readonly dataValidations?: Readonly<Record<string, unknown>>
  readonly sheetProtection?: unknown
  readonly tables?: readonly unknown[]
}

interface ArrayFormulaModel {
  readonly shareType?: 'shared' | 'array'
}

interface ExtendedColor {
  readonly indexed?: number
  readonly tint?: number
}

interface ImportedCell {
  readonly input: CellInput
  readonly value: string | number | boolean | null
}

/** A defined name as the file lists it: `ranges` holds one reference per area. */
export interface DefinedNameEntry {
  readonly name: string
  readonly ranges: readonly string[]
}

const minimumColumnWidth = 48
const maximumColumnWidth = 600
/** Days between the 1900 date system's epoch and the Unix epoch. */
const excelEpochOffset = 25_569
const millisecondsPerDay = 86_400_000

let registeredFunctions: ReadonlySet<string> | undefined

function supportedFunctions(): ReadonlySet<string> {
  registeredFunctions ??= new Set(HyperFormula.getRegisteredFunctionNames('enGB'))
  return registeredFunctions
}

function isRichText(value: ExcelJS.CellValue): value is ExcelJS.CellRichTextValue {
  return typeof value === 'object' && value !== null && 'richText' in value
}

function isHyperlink(value: ExcelJS.CellValue): value is ExcelJS.CellHyperlinkValue {
  return typeof value === 'object' && value !== null && 'hyperlink' in value
}

function isError(value: ExcelJS.CellValue): value is ExcelJS.CellErrorValue {
  return typeof value === 'object' && value !== null && 'error' in value
}

function isFormula(
  value: ExcelJS.CellValue,
): value is ExcelJS.CellFormulaValue | ExcelJS.CellSharedFormulaValue {
  return typeof value === 'object' && value !== null && ('formula' in value || 'sharedFormula' in value)
}

function dateToSerial(date: Date): number {
  return excelEpochOffset + date.getTime() / millisecondsPerDay
}

function serialToDate(serial: number): Date {
  return new Date(Math.round((serial - excelEpochOffset) * millisecondsPerDay))
}

function argbToHex(argb: string): string | undefined {
  const digits = argb.length === 8 ? argb.slice(2) : argb.length === 6 ? argb : undefined
  return digits === undefined || !/^[0-9A-Fa-f]{6}$/.test(digits) ? undefined : `#${digits.toLowerCase()}`
}

function hexToArgb(hex: string): string {
  return `FF${hex.replace('#', '').toUpperCase()}`
}

function columnIndex(letters: string): number {
  return [...letters.toUpperCase()].reduce((total, letter) => total * 26 + (letter.charCodeAt(0) - 64), 0) - 1
}

/** A column width in Excel's character units, as pixels the grid can show. */
export function columnWidthPixels(characters: number): number {
  return Math.max(minimumColumnWidth, Math.min(maximumColumnWidth, Math.round(characters * 7 + 5)))
}

function cachedValue(result: ExcelJS.CellValue): string | number | boolean | null {
  if (result === null || result === undefined) return null
  if (result instanceof Date) return dateToSerial(result)
  if (typeof result === 'number' || typeof result === 'string' || typeof result === 'boolean') return result
  return isError(result) ? result.error : null
}

/** Function names called in a formula, with the markers Excel adds for newer functions removed. */
function calledFunctions(formula: string): readonly string[] {
  const withoutText = formula.replaceAll(/"(?:[^"]|"")*"/g, '""')
  return [...withoutText.matchAll(/(?:_xlfn\.|_xlws\.)*([A-Za-z][A-Za-z0-9_.]*)\s*\(/g)]
    .flatMap(match => (match[1] === undefined ? [] : [match[1].toUpperCase()]))
}

/** Formulas calling functions the calculation engine does not have, counted per sheet under one finding. */
export function unsupportedFunctionEvents(sheets: readonly SheetData[]): FindingEvent[] {
  const supported = supportedFunctions()
  const names = new Set<string>()
  const cellsBySheet = new Map<string, number>()

  for (const sheet of sheets) {
    for (const input of sheet.cells.flat()) {
      if (typeof input !== 'string' || !input.startsWith('=')) continue
      const missing = calledFunctions(input).filter(name => !supported.has(name))
      if (missing.length === 0) continue
      missing.forEach(name => names.add(name))
      cellsBySheet.set(sheet.name, (cellsBySheet.get(sheet.name) ?? 0) + 1)
    }
  }

  const construct = `Unsupported formula functions: ${[...names].sort().join(', ')}`
  const suggestedAlternative = 'These cells show an error here. Their formulas are saved unchanged.'
  return [...cellsBySheet].map(([location, count]) => ({ construct, severity: 'degraded', location, count, unit: 'cells', suggestedAlternative }))
}

interface CellContext {
  readonly sheetName: string
  readonly defaultFont: DefaultFont
  readonly events: FindingEvent[]
}

function cellEvent(context: CellContext, construct: string, severity: 'degraded' | 'dropped' = 'degraded'): void {
  context.events.push({ construct, severity, location: context.sheetName, count: 1, unit: 'cells' })
}

function readCellValue(cell: ExcelJS.Cell, context: CellContext): ImportedCell | undefined {
  const value = cell.value

  switch (cell.type) {
    case ExcelJS.ValueType.Number:
      return typeof value === 'number' ? { input: value, value } : undefined

    case ExcelJS.ValueType.Boolean:
      return typeof value === 'boolean' ? { input: value, value } : undefined

    case ExcelJS.ValueType.String:
    case ExcelJS.ValueType.SharedString:
      return typeof value === 'string' ? { input: value, value } : undefined

    case ExcelJS.ValueType.Date: {
      if (!(value instanceof Date)) return undefined
      const serial = dateToSerial(value)
      return { input: serial, value: serial }
    }

    case ExcelJS.ValueType.RichText: {
      if (!isRichText(value)) return undefined
      cellEvent(context, 'Rich text formatting')
      const text = value.richText.map(run => run.text).join('')
      return { input: text, value: text }
    }

    case ExcelJS.ValueType.Hyperlink: {
      if (!isHyperlink(value)) return undefined
      cellEvent(context, 'Hyperlinks', 'dropped')
      return { input: value.text, value: value.text }
    }

    case ExcelJS.ValueType.Error:
      if (!isError(value)) return undefined
      return { input: value.error, value: value.error }

    case ExcelJS.ValueType.Formula: {
      if ((cell.model as ExcelJS.CellModel & ArrayFormulaModel).shareType === 'array') {
        cellEvent(context, 'Array formulas')
      }
      const formula = cell.formula
      if (typeof formula !== 'string' || formula === '') return undefined
      return { input: `=${formula}`, value: isFormula(value) ? cachedValue(value.result) : null }
    }

    default:
      return undefined
  }
}

function readFont(font: Partial<ExcelJS.Font>, context: CellContext, style: CellStyle): CellStyle {
  let result = style
  const { defaultFont } = context

  if (font.bold === true) result = { ...result, bold: true }
  if (font.italic === true) result = { ...result, italic: true }
  if (font.underline !== undefined && font.underline !== false && font.underline !== 'none') {
    result = { ...result, underline: true }
  }
  if (font.strike === true || font.outline === true || font.vertAlign !== undefined) {
    cellEvent(context, 'Strikethrough, outline and script fonts')
  }
  if (
    (font.name !== undefined && font.name !== defaultFont.name)
    || (font.size !== undefined && font.size !== defaultFont.size)
  ) {
    cellEvent(context, 'Font family and size')
  }

  const color = font.color
  if (color === undefined) return result

  const hex = color.argb === undefined ? undefined : argbToHex(color.argb)
  if (hex !== undefined) return { ...result, textColor: hex }

  const extended = color as Partial<ExcelJS.Color> & ExtendedColor
  const isDefaultTheme = color.theme === defaultFont.colorTheme && extended.tint === undefined
  if (!isDefaultTheme && (color.theme !== undefined || extended.indexed !== undefined)) {
    cellEvent(context, 'Theme and indexed colours')
  }
  return result
}

function readFill(fill: ExcelJS.Fill, context: CellContext, style: CellStyle): CellStyle {
  if (fill.type !== 'pattern') {
    cellEvent(context, 'Gradient fills')
    return style
  }
  if (fill.pattern === 'none' || fill.pattern === undefined) return style
  if (fill.pattern !== 'solid') {
    cellEvent(context, 'Pattern fills')
    return style
  }

  const foreground = fill.fgColor
  if (foreground === undefined) return style
  const hex = foreground.argb === undefined ? undefined : argbToHex(foreground.argb)
  if (hex === undefined) {
    cellEvent(context, 'Theme and indexed colours')
    return style
  }
  return { ...style, fillColor: hex }
}

function readAlignment(
  alignment: Partial<ExcelJS.Alignment>,
  context: CellContext,
  style: CellStyle,
): CellStyle {
  let result = style
  const { horizontal } = alignment

  if (horizontal === 'left' || horizontal === 'center' || horizontal === 'right') {
    result = { ...result, align: horizontal }
  } else if (horizontal !== undefined) {
    cellEvent(context, 'Other text alignment')
  }

  if (alignment.wrapText === true) cellEvent(context, 'Text wrapping')
  if (alignment.vertical !== undefined) cellEvent(context, 'Vertical alignment')
  if (
    (alignment.indent !== undefined && alignment.indent !== 0)
    || (alignment.textRotation !== undefined && alignment.textRotation !== 0)
    || alignment.shrinkToFit === true
    || alignment.readingOrder !== undefined
  ) {
    cellEvent(context, 'Other text alignment')
  }

  return result
}

function readCellStyle(cell: ExcelJS.Cell, context: CellContext): CellStyle {
  const source = cell.style
  let style: CellStyle = defaultCellStyle

  if (source.numFmt !== undefined) {
    const format = numberFormatStyle(source.numFmt)
    style = format.style
    if (!format.mapped) cellEvent(context, 'Unmapped number formats')
  }
  if (source.font !== undefined) style = readFont(source.font, context, style)
  if (source.fill !== undefined) style = readFill(source.fill, context, style)
  if (source.alignment !== undefined) style = readAlignment(source.alignment, context, style)

  const border = source.border
  if (border !== undefined && Object.values(border).some(side => side !== undefined)) {
    cellEvent(context, 'Cell borders')
  }

  return style
}

function readLayout(sheet: ExcelJS.Worksheet, events: FindingEvent[]): SheetData['layout'] {
  const columnWidths: [number, number][] = []
  const hiddenColumns: number[] = []
  const hiddenRows: number[] = []

  ;(sheet.columns ?? []).forEach((column, index) => {
    if (column === undefined || column === null) return
    if (typeof column.width === 'number') columnWidths.push([index, columnWidthPixels(column.width)])
    if (column.hidden === true) hiddenColumns.push(index)
  })

  let customHeights = 0
  sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    if (row.hidden === true) hiddenRows.push(rowNumber - 1)
    if (typeof row.height === 'number') customHeights += 1
  })
  if (customHeights > 0) {
    events.push({
      construct: 'Custom row heights',
      severity: 'degraded',
      location: sheet.name,
      count: customHeights,
      unit: 'rows',
    })
  }

  const view = (sheet.views ?? [])[0]
  if (view?.state === 'split') {
    events.push({ construct: 'Split panes', severity: 'degraded', location: sheet.name })
  }
  const frozen = view?.state === 'frozen' ? view : undefined

  return {
    rowCount: Math.max(defaultRowCount, sheet.rowCount),
    columnCount: Math.max(defaultColumnCount, sheet.columnCount),
    columnWidths,
    hiddenRows,
    hiddenColumns,
    freezeRows: frozen?.ySplit ?? 0,
    freezeColumns: frozen?.xSplit ?? 0,
  }
}

function readSheetConstructs(sheet: ExcelJS.Worksheet, events: FindingEvent[]): void {
  const detail = sheet.model as ExcelJS.WorksheetModel & WorksheetDetail
  const location = sheet.name

  const merges = detail.merges?.length ?? 0
  if (merges > 0) {
    events.push({ construct: 'Merged cells', severity: 'degraded', location, count: merges, unit: 'ranges' })
  }

  const rules = detail.conditionalFormattings?.length ?? 0
  if (rules > 0) {
    events.push({ construct: 'Conditional formatting', severity: 'dropped', location, count: rules, unit: 'ranges' })
  }

  const validations = Object.keys(detail.dataValidations ?? {}).length
  if (validations > 0) {
    events.push({ construct: 'Data validation', severity: 'dropped', location, count: validations, unit: 'cells' })
  }

  const tables = detail.tables?.length ?? 0
  if (tables > 0) {
    events.push({ construct: 'Tables', severity: 'dropped', location, count: tables, unit: 'tables' })
  }

  if (detail.sheetProtection !== undefined && detail.sheetProtection !== null) {
    events.push({ construct: 'Sheet protection', severity: 'dropped', location })
  }
  if (sheet.autoFilter !== undefined && sheet.autoFilter !== null) {
    events.push({ construct: 'Auto filters', severity: 'dropped', location })
  }
  if (sheet.state !== 'visible') {
    events.push({ construct: 'Hidden sheets', severity: 'degraded', location })
  }
}

const definedNamePattern
  = /^(?:'((?:[^']|'')+)'|([^'!]+))!\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?$/

function readNamedRange(name: string, reference: string, sheetNames: ReadonlySet<string>): NamedRange | undefined {
  const match = reference.match(definedNamePattern)
  if (match === null) return undefined

  const sheetName = (match[1]?.replaceAll("''", "'") ?? match[2] ?? '').trim()
  if (!sheetNames.has(sheetName)) return undefined

  const x = columnIndex(match[3] ?? '')
  const y = Number(match[4]) - 1
  const endX = match[5] === undefined ? x : columnIndex(match[5])
  const endY = match[6] === undefined ? y : Number(match[6]) - 1
  if (endX < x || endY < y) return undefined

  return { name, sheetName, range: { x, y, width: endX - x + 1, height: endY - y + 1 } }
}

export function readDefinedNames(
  entries: readonly DefinedNameEntry[],
  sheetNames: ReadonlySet<string>,
  events: FindingEvent[],
): NamedRange[] {
  const namedRanges: NamedRange[] = []

  for (const entry of entries) {
    if (entry.name.startsWith('_xlnm.')) {
      events.push({ construct: 'Print areas and titles', severity: 'degraded', location: entry.name })
      continue
    }

    const reference = entry.ranges.length === 1 ? entry.ranges[0] : undefined
    const namedRange = reference === undefined ? undefined : readNamedRange(entry.name, reference, sheetNames)
    if (namedRange === undefined) {
      events.push({
        construct: 'Defined names that are not a single range',
        severity: 'degraded',
        location: entry.name,
      })
      continue
    }
    namedRanges.push(namedRange)
  }

  return namedRanges
}

function materialize(grid: ReadonlyMap<number, ReadonlyMap<number, ImportedCell>>): {
  readonly cells: CellInput[][]
  readonly values: (string | number | boolean | null)[][]
} {
  const lastRow = Math.max(-1, ...grid.keys())
  const cells: CellInput[][] = []
  const values: (string | number | boolean | null)[][] = []

  for (let row = 0; row <= lastRow; row += 1) {
    const line = grid.get(row)
    const lastColumn = line === undefined ? -1 : Math.max(-1, ...line.keys())
    cells.push(Array.from({ length: lastColumn + 1 }, (_, column) => line?.get(column)?.input ?? null))
    values.push(Array.from({ length: lastColumn + 1 }, (_, column) => line?.get(column)?.value ?? null))
  }

  return { cells, values }
}

export async function readXlsx(bytes: Uint8Array): Promise<XlsxImport> {
  assertZipFitsInMemory(bytes)
  const workbook = new ExcelJS.Workbook()
  await (workbook.xlsx as unknown as ByteXlsx).load(bytes)

  const { parts, defaultFont } = readWorkbookPackage(bytes)
  const events: FindingEvent[] = [...packageFindings(parts)]
  const sheets: SheetData[] = []

  for (const sheet of workbook.worksheets) {
    const context: CellContext = { sheetName: sheet.name, defaultFont, events }
    const grid = new Map<number, Map<number, ImportedCell>>()
    const styles: [number, number, CellStyle][] = []

    sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
        const imported = readCellValue(cell, context)
        if (imported !== undefined) {
          const line = grid.get(rowNumber - 1) ?? new Map<number, ImportedCell>()
          line.set(columnNumber - 1, imported)
          grid.set(rowNumber - 1, line)
        }

        const style = readCellStyle(cell, context)
        if (!isDefaultStyle(style)) styles.push([columnNumber - 1, rowNumber - 1, style])
      })
    })

    readSheetConstructs(sheet, events)
    const { cells, values } = materialize(grid)
    sheets.push({ name: sheet.name, cells, values, styles, layout: readLayout(sheet, events) })
  }

  events.push(...unsupportedFunctionEvents(sheets))
  const sheetNames = new Set(sheets.map(sheet => sheet.name))
  const namedRanges = readDefinedNames(workbook.definedNames.model, sheetNames, events)

  return { data: { sheets, namedRanges }, findings: aggregateFindings(events) }
}

function writeCellStyle(cell: ExcelJS.Cell, style: CellStyle): void {
  const pattern = style.formatCode ?? numberFormatPattern(style)
  if (pattern !== undefined) cell.numFmt = pattern

  const font: Partial<ExcelJS.Font> = {}
  if (style.bold === true) font.bold = true
  if (style.italic === true) font.italic = true
  if (style.underline === true) font.underline = true
  if (style.textColor !== undefined) font.color = { argb: hexToArgb(style.textColor) }
  if (Object.keys(font).length > 0) cell.font = font

  if (style.fillColor !== undefined) {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(style.fillColor) } }
  }
  if (style.align !== undefined) cell.alignment = { horizontal: style.align }
}

function writeCellValue(
  cell: ExcelJS.Cell,
  input: CellInput,
  cached: string | number | boolean | null,
  style: CellStyle | undefined,
): void {
  if (input === null || input === '') return

  if (typeof input === 'string' && input.startsWith('=')) {
    // A cached result lets Excel, Numbers and Google Sheets show values before they recalculate.
    const result = style?.numberFormat === 'date' && typeof cached === 'number' ? serialToDate(cached) : cached
    cell.value = { formula: input.slice(1), result: result === null ? undefined : result }
    return
  }

  if (style?.numberFormat === 'date' && typeof input === 'number') {
    cell.value = serialToDate(input)
    return
  }

  cell.value = input
}

export async function writeXlsx(data: WorkbookData): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook()

  for (const sheet of data.sheets) {
    const target = workbook.addWorksheet(sheet.name)
    const styles = new Map<string, CellStyle>(
      sheet.styles.map(([column, row, style]) => [`${column}:${row}`, style]),
    )

    sheet.cells.forEach((line, row) => {
      line.forEach((input, column) => {
        const cached = sheet.values?.[row]?.[column] ?? null
        writeCellValue(target.getCell(row + 1, column + 1), input, cached, styles.get(`${column}:${row}`))
      })
    })
    sheet.styles.forEach(([column, row, style]) => {
      writeCellStyle(target.getCell(row + 1, column + 1), style)
    })

    const { layout } = sheet
    layout.columnWidths.forEach(([column, width]) => {
      target.getColumn(column + 1).width = (width - 5) / 7
    })
    layout.hiddenColumns.forEach(column => {
      target.getColumn(column + 1).hidden = true
    })
    layout.hiddenRows.forEach(row => {
      target.getRow(row + 1).hidden = true
    })
    if (layout.freezeRows > 0 || layout.freezeColumns > 0) {
      target.views = [{ state: 'frozen', xSplit: layout.freezeColumns, ySplit: layout.freezeRows }]
    }
  }

  for (const namedRange of data.namedRanges) {
    workbook.definedNames.add(sheetRangeReference(namedRange.sheetName, namedRange.range), namedRange.name)
  }

  return new Uint8Array(await (workbook.xlsx as unknown as ByteXlsx).writeBuffer())
}
