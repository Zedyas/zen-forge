import {
  allRuns,
  formatRuns,
  holdsText,
  insertColumn,
  insertRow,
  newTable,
  removeColumn,
  removeRow,
  updateCell,
  withHeaderRow,
  type Border,
  type HorizontalAlign,
  type ListStyle,
  type Paragraph,
  type RunStyle,
  type SlideElement,
  type TableCell,
  type TableElement,
  type VerticalAlign,
} from './model'
import { addElement, changeElement, changeElements, setCurrentCell } from './slides-actions'
import { readySlides, selectedElements, type ReadySlides, type TableCellAddress } from './slides-store'
import { activeEditor } from './text-editing'
import { findTheme } from './themes'

/*
 * Formatting from the toolbar, menus and inspector. While a text box is being typed into it goes to
 * the typed text's selection; otherwise to every selected text box and shape, and for a table to
 * its current cell, or to all its cells when none is current.
 */

function selection(id: string): { readonly document: ReadySlides; readonly elements: readonly SlideElement[] } | undefined {
  const document = readySlides(id)
  return document === undefined ? undefined : { document, elements: selectedElements(document) }
}

/** The cells a change applies to: the current one, or every cell. */
function targetCells(table: TableElement, cell: TableCellAddress | undefined): readonly TableCellAddress[] {
  if (cell !== undefined) return [cell]
  return table.rows.flatMap((row, rowIndex) => row.cells.map((_, column) => ({ row: rowIndex, column })))
}

function changeCells(table: TableElement, cell: TableCellAddress | undefined, change: Partial<TableCell>): TableElement {
  return targetCells(table, cell).reduce((current, address) => updateCell(current, address.row, address.column, change), table)
}

function changeParagraphs(id: string, ids: readonly string[], change: (paragraphs: readonly Paragraph[]) => Paragraph[]): void {
  changeElements(id, ids, element => holdsText(element) ? { ...element, paragraphs: change(element.paragraphs) } : element)
}

export function applyRunStyle(id: string, change: Partial<RunStyle>): void {
  const editor = activeEditor(id)
  if (editor !== undefined) {
    if (change.color !== undefined) editor.format({ kind: 'color', value: change.color })
    if (change.size !== undefined) editor.format({ kind: 'size', value: change.size })
    if (change.font !== undefined) editor.format({ kind: 'font', value: change.font })
    if ('highlight' in change) editor.format({ kind: 'highlight', value: change.highlight })
    return
  }
  const current = selection(id)
  if (current === undefined) return
  const { document, elements } = current
  changeElements(id, elements.map(element => element.id), element => {
    if (holdsText(element)) return { ...element, paragraphs: formatRuns(element.paragraphs, change) }
    if (element.kind !== 'table') return element
    const table = { ...element, size: change.size ?? element.size, font: change.font ?? element.font }
    const cells: Partial<TableCell> = { ...(change.color === undefined ? {} : { color: change.color }), ...(change.bold === undefined ? {} : { bold: change.bold }) }
    return Object.keys(cells).length === 0 ? table : changeCells(table, document.cell, cells)
  })
}

/** Whether every selected run (or table cell, for bold) already has a style, so the toggle turns it off. */
function everyHas(document: ReadySlides, elements: readonly SlideElement[], style: 'bold' | 'italic' | 'underline'): boolean {
  const flags = elements.flatMap(element => {
    if (holdsText(element)) return allRuns(element.paragraphs).map(run => run[style])
    if (element.kind === 'table' && style === 'bold') return targetCells(element, document.cell).map(({ row, column }) => element.rows[row]?.cells[column]?.bold ?? false)
    return []
  })
  return flags.length > 0 && flags.every(Boolean)
}

export function toggleRunStyle(id: string, style: 'bold' | 'italic' | 'underline'): void {
  const editor = activeEditor(id)
  if (editor !== undefined) {
    editor.format({ kind: 'toggle', style })
    return
  }
  const current = selection(id)
  if (current !== undefined) applyRunStyle(id, { [style]: !everyHas(current.document, current.elements, style) })
}

export function setAlign(id: string, align: HorizontalAlign): void {
  const editor = activeEditor(id)
  if (editor !== undefined) {
    editor.format({ kind: 'align', value: align })
    return
  }
  const current = selection(id)
  if (current === undefined) return
  changeElements(id, current.elements.map(element => element.id), element => {
    if (holdsText(element)) return { ...element, paragraphs: element.paragraphs.map(paragraph => ({ ...paragraph, align })) }
    return element.kind === 'table' ? changeCells(element, current.document.cell, { align }) : element
  })
}

/** Turns bullets or numbering on for the selected paragraphs, or off when they all have it already. */
export function toggleList(id: string, list: Exclude<ListStyle, 'none'>): void {
  const editor = activeEditor(id)
  if (editor !== undefined) {
    editor.format({ kind: 'list', value: list })
    return
  }
  const current = selection(id)
  if (current === undefined) return
  const paragraphs = current.elements.flatMap(element => holdsText(element) ? element.paragraphs : [])
  const next: ListStyle = paragraphs.length > 0 && paragraphs.every(paragraph => paragraph.list === list) ? 'none' : list
  changeParagraphs(id, current.document.selection, all => all.map(paragraph => ({ ...paragraph, list: next })))
}

export function changeIndent(id: string, step: 1 | -1): void {
  const editor = activeEditor(id)
  if (editor !== undefined) {
    editor.format({ kind: 'indent', step })
    return
  }
  const document = readySlides(id)
  if (document !== undefined) {
    changeParagraphs(id, document.selection, all => all.map(paragraph => ({ ...paragraph, level: Math.min(8, Math.max(0, paragraph.level + step)) })))
  }
}

export function setLineSpacing(id: string, lineSpacing: number): void {
  const document = readySlides(id)
  if (document !== undefined) changeParagraphs(id, document.selection, all => all.map(paragraph => ({ ...paragraph, lineSpacing })))
}

export function setVerticalAlign(id: string, verticalAlign: VerticalAlign): void {
  const current = selection(id)
  if (current === undefined) return
  changeElements(id, current.document.selection, element => {
    if (holdsText(element)) return { ...element, verticalAlign }
    return element.kind === 'table' ? changeCells(element, current.document.cell, { verticalAlign }) : element
  })
}

/** Fills the selected shapes and text boxes, or the current table cell. */
export function setFill(id: string, fill: string | undefined): void {
  const current = selection(id)
  if (current === undefined) return
  changeElements(id, current.document.selection, element => {
    if (element.kind === 'text' || element.kind === 'shape') return { ...element, fill }
    return element.kind === 'table' ? changeCells(element, current.document.cell, { fill }) : element
  })
}

/** Outlines the selected shapes and text boxes, or sets a table's cell borders; tables always keep a border. */
export function setBorder(id: string, border: Border | undefined): void {
  const document = readySlides(id)
  if (document === undefined) return
  changeElements(id, document.selection, element => {
    if (element.kind === 'text' || element.kind === 'shape') return { ...element, border }
    return element.kind === 'table' ? { ...element, border: border ?? { ...element.border, width: 0 } } : element
  })
}

/* ─── Tables ─── */

export function insertTable(id: string, rows: number, columns: number): void {
  const document = readySlides(id)
  if (document !== undefined) addElement(id, newTable(rows, columns, document.present))
}

function selectedTable(document: ReadySlides): TableElement | undefined {
  const [element, ...rest] = selectedElements(document)
  return element?.kind === 'table' && rest.length === 0 ? element : undefined
}

export type TableEdit = 'rowAbove' | 'rowBelow' | 'columnLeft' | 'columnRight' | 'deleteRow' | 'deleteColumn'

/** Adds or removes a row or column beside the current cell, or at the table's end when none is current. */
export function editTable(id: string, edit: TableEdit): void {
  const document = readySlides(id)
  const table = document === undefined ? undefined : selectedTable(document)
  if (document === undefined || table === undefined) return
  const cell = document.cell ?? { row: table.rows.length - 1, column: table.columns.length - 1 }
  const change: Record<TableEdit, () => TableElement> = {
    rowAbove: () => insertRow(table, cell.row),
    rowBelow: () => insertRow(table, cell.row + 1),
    columnLeft: () => insertColumn(table, cell.column),
    columnRight: () => insertColumn(table, cell.column + 1),
    deleteRow: () => removeRow(table, cell.row),
    deleteColumn: () => removeColumn(table, cell.column),
  }
  changeElement(id, table.id, () => change[edit]())
  // The current cell follows the row or column it was in.
  const next = {
    row: edit === 'rowAbove' ? cell.row + 1 : edit === 'deleteRow' ? Math.max(0, cell.row - 1) : cell.row,
    column: edit === 'columnLeft' ? cell.column + 1 : edit === 'deleteColumn' ? Math.max(0, cell.column - 1) : cell.column,
  }
  if (document.cell !== undefined) setCurrentCell(id, next)
}

export function setCellText(id: string, tableId: string, cell: TableCellAddress, text: string): void {
  changeElement(id, tableId, element => {
    if (element.kind !== 'table' || element.rows[cell.row]?.cells[cell.column]?.text === text) return element
    return updateCell(element, cell.row, cell.column, { text })
  })
}

export function setHeaderRow(id: string, on: boolean): void {
  const document = readySlides(id)
  const table = document === undefined ? undefined : selectedTable(document)
  if (document !== undefined && table !== undefined) changeElement(id, table.id, () => withHeaderRow(table, on, findTheme(document.present.theme)))
}

export function setTableBorderColor(id: string, color: string): void {
  const document = readySlides(id)
  const table = document === undefined ? undefined : selectedTable(document)
  if (table !== undefined) changeElement(id, table.id, element => element.kind === 'table' ? { ...element, border: { color, width: Math.max(1, element.border.width) } } : element)
}
