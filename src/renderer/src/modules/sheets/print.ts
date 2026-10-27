import type { Item } from '@glideapps/glide-data-grid'
import type { Paper } from '@shared/print'
import { pageSetup, printableArea, type PageSetup, type Printout } from '../../services/print/print'
import { fontPixels } from './cell-text'
import { viewRange } from './grid-mapping'
import { formatCellValue, formatColor } from './model/format'
import type { CellStyle, Workbook } from './model/Workbook'
import type { CellRange, HorizontalAlign } from './model/workbook-data'

export interface PrintCell {
  readonly text: string
  readonly style: CellStyle
  /** The cell's own alignment, else the grid's: numbers right, everything else left. */
  readonly align: HorizontalAlign
  /** The merged block holding the cell, in table columns and rows; the block prints as one cell. */
  readonly block?: CellRange
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

/** The used size of a sheet, grown to the far edge of any merged cell that starts inside it, so a merged title keeps its size. */
function printedSize(workbook: Workbook, sheetId: number): { readonly width: number; readonly height: number } {
  const used = workbook.usedSize(sheetId)
  return workbook.merges(sheetId).reduce((size, merge) => merge.x < used.width && merge.y < used.height
    ? { width: Math.max(size.width, merge.x + merge.width), height: Math.max(size.height, merge.y + merge.height) }
    : size, used)
}

/** The printable table of a sheet, or undefined when its visible cells are all empty. */
export function sheetPrintTable(workbook: Workbook, sheetId: number): SheetPrintTable | undefined {
  const { width, height } = printedSize(workbook, sheetId)
  const rows = workbook.visibleRows(sheetId).filter(row => row < height)
  const columns = workbook.visibleColumns(sheetId).filter(column => column < width)
  const blocks = workbook.merges(sheetId).flatMap(merge => {
    const block = viewRange(merge, columns, rows)
    return block === undefined ? [] : [{ merge, block }]
  })
  const cells = rows.map((row, rowIndex) => columns.map((column, columnIndex): PrintCell => {
    const found = blocks.find(({ block }) => columnIndex >= block.x && columnIndex < block.x + block.width
      && rowIndex >= block.y && rowIndex < block.y + block.height)
    // A merged block prints one cell, from the merge's upper-left cell even when that cell is hidden or on another page.
    const source: Item = found === undefined ? [column, row] : [found.merge.x, found.merge.y]
    const value = workbook.getCellValue(sheetId, source)
    const style = workbook.getStyle(sheetId, source)
    const color = formatColor(value, style)
    return {
      text: formatCellValue(value, style),
      style: color === undefined ? style : { ...style, textColor: color },
      align: style.align ?? (typeof value === 'number' ? 'right' : 'left'),
      ...(found === undefined ? {} : { block: found.block }),
    }
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
  const { bold, italic, underline, strike, wrap, fontFamily, fontSize, textColor, fillColor } = cell.style
  element.textContent = cell.text
  element.style.textAlign = cell.align
  if (bold === true) element.style.fontWeight = '600'
  if (italic === true) element.style.fontStyle = 'italic'
  const decorations = [underline === true ? 'underline' : '', strike === true ? 'line-through' : ''].filter(Boolean)
  if (decorations.length > 0) element.style.textDecoration = decorations.join(' ')
  if (wrap === true) element.style.whiteSpace = 'pre-wrap'
  if (fontFamily !== undefined) element.style.fontFamily = `"${fontFamily.replaceAll('"', '')}"`
  if (fontSize !== undefined) element.style.fontSize = `${fontPixels(cell.style)}px`
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
  // `offset` is the section's first table row, which merged blocks count from.
  const addRows = (section: HTMLTableSectionElement, rows: readonly (readonly PrintCell[])[], offset: number): void => rows.forEach((row, index) => {
    const line = section.insertRow()
    columns.forEach((column, position) => {
      const cell = row[column]
      const block = cell?.block
      if (cell === undefined) return
      if (block === undefined) {
        line.append(cellElement(cell))
        return
      }
      // A merged block prints from its first row and column on this page and in this section; its other cells are covered.
      const inBlock = (candidate: number | undefined): boolean => candidate !== undefined && candidate >= block.x && candidate < block.x + block.width
      if (offset + index !== Math.max(block.y, offset) || inBlock(columns[position - 1])) return
      const merged = cellElement(cell)
      merged.colSpan = columns.filter(inBlock).length
      merged.rowSpan = Math.min(block.y + block.height, offset + rows.length) - (offset + index)
      line.append(merged)
    })
  })
  // A table header repeats at the top of every printed page.
  if (table.headRows.length > 0) addRows(element.createTHead(), table.headRows, 0)
  addRows(element.createTBody(), table.bodyRows, table.headRows.length)
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
