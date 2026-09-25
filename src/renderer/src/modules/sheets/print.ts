import type { Paper } from '@shared/print'
import { pageSetup, printableArea, type PageSetup, type Printout } from '../../services/print/print'
import { formatCellValue } from './model/format'
import type { CellStyle, Workbook } from './model/Workbook'
import type { HorizontalAlign } from './model/workbook-data'

export interface PrintCell {
  readonly text: string
  readonly style: CellStyle
  /** The cell's own alignment, else the grid's: numbers right, everything else left. */
  readonly align: HorizontalAlign
}

/** A sheet's used range as it prints: hidden rows and columns left out, values formatted as the grid shows them. */
export interface SheetPrintTable {
  /** Width of each printed column, in CSS pixels. */
  readonly columnWidths: readonly number[]
  /** Frozen columns at the left. They label the rows, so they repeat when the columns need several pages across. */
  readonly frozenColumns: number
  /** Frozen rows. They head every page. */
  readonly headRows: readonly (readonly PrintCell[])[]
  readonly bodyRows: readonly (readonly PrintCell[])[]
}

/** The printable table of a sheet, or undefined when its visible cells are all empty. */
export function sheetPrintTable(workbook: Workbook, sheetId: number): SheetPrintTable | undefined {
  const { width, height } = workbook.usedSize(sheetId)
  const rows = workbook.visibleRows(sheetId).filter(row => row < height)
  const columns = workbook.visibleColumns(sheetId).filter(column => column < width)
  const cells = rows.map(row => columns.map((column): PrintCell => {
    const value = workbook.getCellValue(sheetId, [column, row])
    const style = workbook.getStyle(sheetId, [column, row])
    return { text: formatCellValue(value, style), style, align: style.align ?? (typeof value === 'number' ? 'right' : 'left') }
  }))
  if (cells.every(row => row.every(cell => cell.text === ''))) return undefined
  // Frozen rows are the sheet's top rows, so the visible ones start the printed rows.
  const frozenRows = Math.min(workbook.getFrozenRowCount(sheetId), rows.length)
  return {
    columnWidths: columns.map(column => workbook.getColumnWidth(sheetId, column)),
    frozenColumns: Math.min(workbook.getFrozenColumnCount(sheetId), columns.length),
    headRows: cells.slice(0, frozenRows),
    bodyRows: cells.slice(frozenRows),
  }
}

const margin = 12
const cssPixelsPerMillimetre = 96 / 25.4
/** Below this, text gets too small to read; a wider table continues on more pages across instead. */
const minimumScale = 0.6
/** A table up to this much wider than a portrait page shrinks to fit it; a wider one turns the page to landscape. */
const portraitAllowance = 1.15

export interface SheetPagePlan {
  readonly page: PageSetup
  /** Zoom on the table so its columns fit the page width; 1 when they already fit. */
  readonly scale: number
  /** The columns on each page across, as table column indexes. All rows print with the first group, then the next. */
  readonly columnGroups: readonly (readonly number[])[]
}

function sum(values: readonly number[]): number {
  return values.reduce((total, value) => total + value, 0)
}

function printableWidth(page: PageSetup): number {
  return printableArea(page).width * cssPixelsPerMillimetre
}

/** Splits the columns into groups at most `limit` pixels wide, repeating the frozen columns unless they fill half a page. */
function columnGroups(widths: readonly number[], frozenColumns: number, limit: number): number[][] {
  const columns = widths.map((_, index) => index)
  if (sum(widths) <= limit + 0.5) return [columns]
  const frozenWidth = sum(widths.slice(0, frozenColumns))
  const repeated = frozenWidth <= limit / 2 ? columns.slice(0, frozenColumns) : []
  const start = repeated.length === 0 ? 0 : frozenWidth
  const groups: number[][] = []
  let group = [...repeated]
  let used = start
  widths.forEach((width, column) => {
    if (column < repeated.length) return
    if (group.length > repeated.length && used + width > limit) {
      groups.push(group)
      group = [...repeated]
      used = start
    }
    group.push(column)
    used += width
  })
  groups.push(group)
  return groups
}

/** Portrait unless the table is clearly wider, then scaled down to fit the page width, down to a minimum. */
export function planSheetPages(columnWidths: readonly number[], frozenColumns: number, paper: Paper): SheetPagePlan {
  const tableWidth = sum(columnWidths)
  const portrait = pageSetup(paper, 'portrait', margin)
  const page = tableWidth <= printableWidth(portrait) * portraitAllowance ? portrait : pageSetup(paper, 'landscape', margin)
  const scale = Math.min(1, Math.max(minimumScale, printableWidth(page) / tableWidth))
  return { page, scale, columnGroups: columnGroups(columnWidths, frozenColumns, printableWidth(page) / scale) }
}

function cellElement(cell: PrintCell): HTMLTableCellElement {
  const element = document.createElement('td')
  const { bold, italic, underline, textColor, fillColor } = cell.style
  element.textContent = cell.text
  element.style.textAlign = cell.align
  if (bold === true) element.style.fontWeight = '600'
  if (italic === true) element.style.fontStyle = 'italic'
  if (underline === true) element.style.textDecoration = 'underline'
  if (textColor !== undefined) element.style.color = textColor
  if (fillColor !== undefined) element.style.backgroundColor = fillColor
  return element
}

function tableElement(table: SheetPrintTable, columns: readonly number[]): HTMLTableElement {
  const element = document.createElement('table')
  const colgroup = element.appendChild(document.createElement('colgroup'))
  const widths = columns.map(column => table.columnWidths[column] ?? 0)
  widths.forEach(width => {
    colgroup.appendChild(document.createElement('col')).style.width = `${width}px`
  })
  element.style.width = `${sum(widths)}px`
  const addRows = (section: HTMLTableSectionElement, rows: readonly (readonly PrintCell[])[]): void => rows.forEach(row => {
    const line = section.insertRow()
    columns.forEach(column => {
      const cell = row[column]
      if (cell !== undefined) line.append(cellElement(cell))
    })
  })
  // A table header repeats at the top of every printed page.
  if (table.headRows.length > 0) addRows(element.createTHead(), table.headRows)
  addRows(element.createTBody(), table.bodyRows)
  return element
}

/** The printout of a sheet table: one table per group of columns, each starting a new page. Styles are in sheets.css. */
export function sheetPrintout(table: SheetPrintTable, paper: Paper): Printout {
  const plan = planSheetPages(table.columnWidths, table.frozenColumns, paper)
  const content = document.createElement('div')
  content.className = 'sheet-print'
  content.style.zoom = String(plan.scale)
  content.append(...plan.columnGroups.map(group => tableElement(table, group)))
  return { content, page: plan.page }
}
