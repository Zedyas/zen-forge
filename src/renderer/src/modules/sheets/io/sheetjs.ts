/**
 * Excel 97–2003 (.xls), OpenDocument (.ods) and Numbers (.numbers) import, through SheetJS.
 *
 * SheetJS Community Edition reads values, formulas (not from Numbers), number formats and merges,
 * but no cell styles, so every import reports what the format's reader never looks at, as well as
 * what the file is found to contain. These formats are opened only: saving writes a new .xlsx.
 */

import { unzipSync } from 'fflate'
import type { CellObject, WorkBook, WorkSheet } from 'xlsx'
import type { ImportFindingInput } from '@shared/fidelity'
import {
  defaultColumnCount,
  defaultRowCount,
  isDefaultStyle,
  type CellInput,
  type CellStyle,
  type SheetData,
  type WorkbookData,
} from '../model/workbook-data'
import { aggregateFindings, type FindingEvent } from './findings'
import { assertZipFitsInMemory } from '../../../services/zip'
import { numberFormatStyle } from './number-format'
import { columnWidthPixels, readDefinedNames, unsupportedFunctionEvents } from './xlsx'

export type SheetJsFormat = 'xls' | 'ods' | 'numbers'

export interface SheetJsImport {
  readonly data: WorkbookData
  readonly findings: ImportFindingInput[]
}

export function isSheetJsFormat(extension: string): extension is SheetJsFormat {
  return extension === 'xls' || extension === 'ods' || extension === 'numbers'
}

/** An xls drawing object; SheetJS keeps them on the sheet without publishing a type. */
interface XlsObject {
  /** [id, object type, flags] */
  readonly cmo: readonly [number, number, number]
}

interface SheetDetail {
  readonly '!objects'?: readonly (XlsObject | undefined)[]
}

interface ImportedCell {
  readonly input: CellInput
  readonly value: string | number | boolean | null
  readonly style?: CellStyle
}

/** Days from the 1900 date system's epoch to the 1904 one; the model always counts from 1900. */
const date1904Offset = 1_462
/** Days between the 1900 date system's epoch and the Unix epoch. */
const excelEpochOffset = 25_569
const millisecondsPerDay = 86_400_000
/** xls drawing object types. Comment boxes and auto filter drop-down arrows are reported as comments and filters. */
const xlsChart = 0x05
const xlsNonDrawings = new Set([0x14, 0x19])

const cellFormatting: FindingEvent = {
  construct: 'Cell formatting',
  severity: 'dropped',
  suggestedAlternative: 'Fonts, colours, fills and borders are not read from this format. Number formats are kept.',
}
const defaultLayout: FindingEvent = {
  construct: 'Column widths, hidden rows and columns, and frozen panes',
  severity: 'degraded',
  suggestedAlternative: 'Columns open at the default width, and every row and column is shown and unfrozen.',
}

/** What each format's reader never looks at, so every import reports it. */
const unread: Record<SheetJsFormat, readonly FindingEvent[]> = {
  xls: [
    cellFormatting,
    {
      construct: 'Frozen rows and columns',
      severity: 'degraded',
      suggestedAlternative: 'Freeze them again from the Rows and Columns menus.',
    },
  ],
  ods: [cellFormatting, defaultLayout],
  numbers: [
    {
      ...cellFormatting,
      suggestedAlternative: 'Fonts, colours, fills, borders and number formats other than dates are not read from Numbers files.',
    },
    defaultLayout,
    {
      construct: 'Formulas',
      severity: 'dropped',
      suggestedAlternative: 'Formulas open as their last calculated values. To keep them, export from Numbers as Excel.',
    },
    {
      construct: 'Charts, images, text boxes and shapes',
      severity: 'dropped',
      suggestedAlternative: 'Only tables are read from Numbers files.',
    },
  ],
}

function cellValue(cell: CellObject, isDate: boolean, dateOffset: number): string | number | boolean | null {
  switch (cell.t) {
    case 'n':
      return typeof cell.v === 'number' ? cell.v + (isDate ? dateOffset : 0) : null
    case 's':
      return typeof cell.v === 'string' ? cell.v : null
    case 'b':
      return typeof cell.v === 'boolean' ? cell.v : null
    case 'd':
      return cell.v instanceof Date ? excelEpochOffset + cell.v.getTime() / millisecondsPerDay : null
    case 'e':
      return cell.w ?? '#N/A'
    case 'z':
      return null
  }
}

interface SheetContext {
  readonly sheetName: string
  readonly dateOffset: number
  readonly events: FindingEvent[]
}

function cellEvent(context: SheetContext, construct: string, severity: 'degraded' | 'dropped'): void {
  context.events.push({ construct, severity, location: context.sheetName, count: 1, unit: 'cells' })
}

function readCell(cell: CellObject, context: SheetContext): ImportedCell {
  const format = typeof cell.z === 'string' ? numberFormatStyle(cell.z) : undefined
  if (format?.mapped === false) cellEvent(context, 'Unmapped number formats', 'degraded')
  if (cell.c !== undefined && cell.c.length > 0) cellEvent(context, 'Comments', 'dropped')
  if (cell.l !== undefined) cellEvent(context, 'Hyperlinks', 'dropped')
  if (cell.F !== undefined) cellEvent(context, 'Array formulas', 'degraded')

  const value = cellValue(cell, format?.style.numberFormat === 'date', context.dateOffset)
  const input = typeof cell.f === 'string' && cell.f !== '' ? `=${cell.f}` : value
  return { input, value, ...(format === undefined || isDefaultStyle(format.style) ? {} : { style: format.style }) }
}

/** `usedRows` and `usedColumns` are the extent of the cells read, which the grid always covers. */
function readLayout(sheet: WorkSheet, usedRows: number, usedColumns: number, context: SheetContext): SheetData['layout'] {
  const columnWidths: [number, number][] = []
  const hiddenColumns: number[] = []
  const hiddenRows: number[] = []
  let customHeights = 0

  // Sparse arrays: forEach skips the columns and rows the file says nothing about.
  ;(sheet['!cols'] ?? []).forEach((column, index) => {
    if (column === undefined || column === null) return
    if (column.hidden === true) hiddenColumns.push(index)
    else if (typeof column.width === 'number') columnWidths.push([index, columnWidthPixels(column.width)])
  })
  ;(sheet['!rows'] ?? []).forEach((row, index) => {
    if (row === undefined || row === null) return
    if (row.hidden === true) hiddenRows.push(index)
    else if (typeof row.hpt === 'number') customHeights += 1
  })
  if (customHeights > 0) {
    context.events.push({
      construct: 'Custom row heights',
      severity: 'degraded',
      location: context.sheetName,
      count: customHeights,
      unit: 'rows',
    })
  }

  return {
    rowCount: Math.max(defaultRowCount, usedRows, (hiddenRows.at(-1) ?? -1) + 1),
    columnCount: Math.max(defaultColumnCount, usedColumns, (hiddenColumns.at(-1) ?? -1) + 1),
    columnWidths,
    hiddenRows,
    hiddenColumns,
    freezeRows: 0,
    freezeColumns: 0,
  }
}

function readSheetConstructs(sheet: WorkSheet, context: SheetContext): void {
  const { events, sheetName: location } = context

  const merges = sheet['!merges']?.length ?? 0
  if (merges > 0) {
    events.push({ construct: 'Merged cells', severity: 'degraded', location, count: merges, unit: 'ranges' })
  }
  if (sheet['!autofilter'] !== undefined) events.push({ construct: 'Auto filters', severity: 'dropped', location })
  // An xls Protect record is read as-is, so an unprotected sheet can carry `false` here.
  if (sheet['!protect']) events.push({ construct: 'Sheet protection', severity: 'dropped', location })

  const objects = ((sheet as SheetDetail)['!objects'] ?? []).filter(object => object !== undefined)
  const charts = objects.filter(object => object.cmo[1] === xlsChart).length
  const drawings = objects.filter(object => object.cmo[1] !== xlsChart && !xlsNonDrawings.has(object.cmo[1])).length
  if (charts > 0) events.push({ construct: 'Charts', severity: 'dropped', location, count: charts, unit: 'charts' })
  if (drawings > 0) {
    events.push({ construct: 'Images and drawings', severity: 'dropped', location, count: drawings, unit: 'objects' })
  }
}

/** ods keeps pictures and embedded charts as package entries, which can be listed without inflating them. */
function odsPackageEvents(bytes: Uint8Array): FindingEvent[] {
  const parts: string[] = []
  unzipSync(bytes, {
    filter: file => {
      parts.push(file.name)
      return false
    },
  })
  const events: FindingEvent[] = []
  if (parts.some(part => part.startsWith('Pictures/'))) {
    events.push({ construct: 'Images and drawings', severity: 'dropped' })
  }
  if (parts.some(part => /^Object \d+\//.test(part))) {
    events.push({ construct: 'Charts and embedded objects', severity: 'dropped' })
  }
  return events
}

/**
 * SheetJS turns each table after the first on a Numbers sheet into a sheet of its own, named after
 * the sheet with `_1`, `_2`… added.
 */
function numbersTableEvents(sheetNames: readonly string[]): FindingEvent[] {
  return sheetNames.flatMap((name, index) => {
    const base = name.match(/^(.+)_\d+$/)?.[1]
    if (base === undefined || !sheetNames.slice(0, index).includes(base)) return []
    return [{
      construct: 'Several tables on one sheet',
      severity: 'degraded',
      location: base,
      count: 1,
      unit: 'extra tables',
      suggestedAlternative: 'Each table after the first opens as a sheet of its own, named after the sheet with _1, _2 and so on.',
    }]
  })
}

function readWorkbookSheets(workbook: WorkBook, context: Omit<SheetContext, 'sheetName'>): SheetData[] {
  const sheets: SheetData[] = []

  workbook.SheetNames.forEach((sheetName, index) => {
    const sheet = workbook.Sheets[sheetName]
    if (sheet === undefined) return
    if (sheet['!type'] === 'chart') {
      context.events.push({ construct: 'Chart sheets', severity: 'dropped', location: sheetName })
      return
    }
    if ((workbook.Workbook?.Sheets?.[index]?.Hidden ?? 0) !== 0) {
      context.events.push({ construct: 'Hidden sheets', severity: 'degraded', location: sheetName })
    }

    const sheetContext: SheetContext = { ...context, sheetName }
    const styles: [number, number, CellStyle][] = []
    // Dense rows are sparse arrays; Array.from turns their gaps into empty cells.
    const grid = Array.from(sheet['!data'] ?? [], (row, rowIndex) => Array.from(row ?? [], (cell, columnIndex) => {
      if (cell === undefined) return undefined
      const imported = readCell(cell, sheetContext)
      if (imported.style !== undefined) styles.push([columnIndex, rowIndex, imported.style])
      return imported
    }))

    readSheetConstructs(sheet, sheetContext)
    const usedColumns = grid.reduce((widest, row) => Math.max(widest, row.length), 0)
    sheets.push({
      name: sheetName,
      cells: grid.map(row => row.map(cell => cell?.input ?? null)),
      values: grid.map(row => row.map(cell => cell?.value ?? null)),
      styles,
      layout: readLayout(sheet, grid.length, usedColumns, sheetContext),
    })
  })

  return sheets
}

export async function readSheetJs(bytes: Uint8Array, format: SheetJsFormat): Promise<SheetJsImport> {
  assertZipFitsInMemory(bytes)
  // Loaded on first use: SheetJS is large, and only these formats need it.
  const XLSX = await import('xlsx')
  let workbook: WorkBook
  try {
    workbook = XLSX.read(bytes, {
      type: 'array',
      dense: true,
      cellNF: true,
      // Only the xls reader needs styles on, for column widths and hidden columns.
      cellStyles: format === 'xls',
      bookVBA: format === 'xls',
    })
  } catch (error) {
    // Numbers allows sheet names that SheetJS refuses: longer than 31 characters, or with : \ / ? * [ ].
    if (format === 'numbers' && error instanceof Error && error.message.startsWith('Sheet name')) {
      throw new Error(`${error.message}. Rename the sheet in Numbers, then open the file again.`, { cause: error })
    }
    throw error
  }

  const events: FindingEvent[] = [...unread[format]]
  if (format === 'ods') events.push(...odsPackageEvents(bytes))
  if (format === 'numbers') events.push(...numbersTableEvents(workbook.SheetNames))
  if (workbook.vbaraw !== undefined) events.push({ construct: 'Macros (VBA)', severity: 'dropped' })

  const dateOffset = workbook.Workbook?.WBProps?.date1904 === true ? date1904Offset : 0
  const sheets = readWorkbookSheets(workbook, { dateOffset, events })
  if (sheets.length === 0) throw new Error('This file has no sheets with cells to open.')

  events.push(...unsupportedFunctionEvents(sheets))
  const names = (workbook.Workbook?.Names ?? []).map(name => ({ name: name.Name, ranges: [name.Ref] }))
  const namedRanges = readDefinedNames(names, new Set(sheets.map(sheet => sheet.name)), events)

  return { data: { sheets, namedRanges }, findings: aggregateFindings(events) }
}
