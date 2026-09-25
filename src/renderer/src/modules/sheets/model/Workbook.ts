import type { Item } from '@glideapps/glide-data-grid'
import {
  HyperFormula,
  type CellValue,
  type FunctionDetails,
  type FunctionListEntry,
  type RawCellContent,
} from 'hyperformula'
import { columnName, sheetRangeReference } from './address'
import { parseCellText } from './cell-parse'
import {
  defaultCellStyle,
  defaultColumnCount,
  defaultRowCount,
  isDefaultStyle,
  type CellInput,
  type CellRange,
  type CellStyle,
  type NamedRange,
  type NumberFormat,
  type SheetData,
  type SheetLayoutData,
  type WorkbookData,
} from './workbook-data'

export type { CellInput, CellRange, CellStyle, NamedRange, NumberFormat, WorkbookData } from './workbook-data'

export interface WorkbookSheet {
  readonly id: number
  readonly name: string
}

export interface SelectionStats {
  /** Non-empty cells. */
  readonly count: number
  /** Cells holding numbers; sum, average, min and max cover only these. */
  readonly numericCount: number
  readonly sum: number
  readonly min: number
  readonly max: number
}

export interface FindMatch {
  readonly sheetId: number
  readonly sheetName: string
  readonly cell: Item
  readonly input: string
}

interface SheetLayout {
  rowCount: number
  columnCount: number
  readonly columnWidths: Map<number, number>
  readonly hiddenRows: Set<number>
  readonly hiddenColumns: Set<number>
  freezeRows: number
  freezeColumns: number
}

interface MutableCellRange {
  x: number
  y: number
  width: number
  height: number
}

const workbookConfig = {
  licenseKey: 'gpl-v3',
  precisionRounding: 10,
  // Undo history is owned by the snapshot stacks below; the engine keeps none.
  undoLimit: 0,
} as const

const defaultColumnWidth = 120
const historyLimit = 100

const blankLayout: SheetLayoutData = {
  rowCount: defaultRowCount,
  columnCount: defaultColumnCount,
  columnWidths: [],
  hiddenRows: [],
  hiddenColumns: [],
  freezeRows: 0,
  freezeColumns: 0,
}

function layoutFromData(layout: SheetLayoutData): SheetLayout {
  return {
    rowCount: Math.max(layout.rowCount, defaultRowCount),
    columnCount: Math.max(layout.columnCount, defaultColumnCount),
    columnWidths: new Map(layout.columnWidths),
    hiddenRows: new Set(layout.hiddenRows),
    hiddenColumns: new Set(layout.hiddenColumns),
    freezeRows: layout.freezeRows,
    freezeColumns: layout.freezeColumns,
  }
}

export function blankWorkbookData(): WorkbookData {
  return { sheets: [{ name: 'Sheet1', cells: [], styles: [], layout: blankLayout }], namedRanges: [] }
}

/** Merges a patch; a patch value of `undefined` removes that property (back to its default). */
function mergeStyle(current: CellStyle, patch: Partial<CellStyle>): CellStyle {
  const merged: Record<string, unknown> = { ...current, ...patch }
  for (const key of Object.keys(merged)) if (merged[key] === undefined) delete merged[key]
  return { ...defaultCellStyle, ...merged }
}

function serializedInput(value: RawCellContent): CellInput {
  if (value === undefined) return null
  if (value instanceof Date) return value.toISOString()
  return value
}

function plainValue(value: CellValue): string | number | boolean | null {
  if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value
  return value.toString()
}

/** Numbers sort before text, empty cells always last; text compares naturally ("Item 2" < "Item 10"). */
function compareForSort(left: CellValue, right: CellValue, ascending: boolean): number {
  const rank = (value: CellValue): number => value === null || value === '' ? 2 : typeof value === 'number' ? 0 : 1
  const leftRank = rank(left)
  const rightRank = rank(right)
  if (leftRank === 2 || rightRank === 2) return leftRank - rightRank
  const direction = ascending ? 1 : -1
  if (leftRank !== rightRank) return (leftRank - rightRank) * direction
  if (typeof left === 'number' && typeof right === 'number') return (left - right) * direction
  return String(left).localeCompare(String(right), undefined, { numeric: true, sensitivity: 'base' }) * direction
}

function styleKey(col: number, row: number): string {
  return `${row}:${col}`
}

function parseStyleKey(key: string): Item {
  const [row = 0, col = 0] = key.split(':').map(Number)
  return [col, row]
}

function namedRangeFormula(namedRange: NamedRange): string {
  return `=${sheetRangeReference(namedRange.sheetName, namedRange.range)}`
}

function cloneRange(range: CellRange): MutableCellRange {
  return { x: range.x, y: range.y, width: range.width, height: range.height }
}

/** HyperFormula-backed workbook with one history for data, layout, formats, and sheets. */
export class Workbook {
  #engine: HyperFormula
  readonly #listeners = new Set<() => void>()
  readonly #layouts = new Map<string, SheetLayout>()
  readonly #styles = new Map<string, Map<string, CellStyle>>()
  readonly #namedRanges = new Map<string, NamedRange>()
  readonly #undoStack: WorkbookData[] = []
  readonly #redoStack: WorkbookData[] = []
  #revision = 0
  #sheetOrder: string[] = []
  /** Undo steps from the loaded state; the document is dirty when it differs from the saved depth. */
  #depth = 0
  /** `undefined` once the saved state is only reachable through a discarded redo branch. */
  #savedDepth: number | undefined = 0

  constructor(data: WorkbookData = blankWorkbookData()) {
    this.#engine = HyperFormula.buildEmpty(workbookConfig)
    this.#load(data.sheets.length > 0 ? data : blankWorkbookData())
  }

  subscribe = (listener: () => void): (() => void) => {
    this.#listeners.add(listener)
    return () => this.#listeners.delete(listener)
  }

  getSnapshot = (): number => this.#revision

  sheets(): readonly WorkbookSheet[] {
    return this.#sheetOrder.map(name => ({ id: this.#sheetId(name), name }))
  }

  getCellInput(sheetId: number, [col, row]: Item): string {
    const serialized = this.#engine.getCellSerialized({ sheet: sheetId, col, row })
    if (serialized instanceof Date) return serialized.toISOString()
    return serialized === null || serialized === undefined ? '' : String(serialized)
  }

  getCellValue(sheetId: number, [col, row]: Item): string | number | boolean | null {
    return plainValue(this.#engine.getCellValue({ sheet: sheetId, col, row }))
  }

  getRowCount(sheetId: number): number {
    return this.#layout(sheetId).rowCount
  }

  getColumnCount(sheetId: number): number {
    return this.#layout(sheetId).columnCount
  }

  visibleRows(sheetId: number): readonly number[] {
    const layout = this.#layout(sheetId)
    return Array.from({ length: layout.rowCount }, (_, index) => index)
      .filter(index => !layout.hiddenRows.has(index))
  }

  visibleColumns(sheetId: number): readonly number[] {
    const layout = this.#layout(sheetId)
    return Array.from({ length: layout.columnCount }, (_, index) => index)
      .filter(index => !layout.hiddenColumns.has(index))
  }

  getColumnWidth(sheetId: number, column: number): number {
    return this.#layout(sheetId).columnWidths.get(column) ?? defaultColumnWidth
  }

  getFrozenColumnCount(sheetId: number): number {
    const layout = this.#layout(sheetId)
    return this.visibleColumns(sheetId).filter(column => column < layout.freezeColumns).length
  }

  getFrozenRowCount(sheetId: number): number {
    const layout = this.#layout(sheetId)
    return this.visibleRows(sheetId).filter(row => row < layout.freezeRows).length
  }

  getStyle(sheetId: number, [col, row]: Item): CellStyle {
    return this.#style(sheetId, col, row)
  }

  setNumberFormat(sheetId: number, range: CellRange, format: NumberFormat): void {
    this.updateStyle(sheetId, range, { numberFormat: format })
  }

  changeDecimalPlaces(sheetId: number, range: CellRange, difference: -1 | 1): void {
    this.#updateStyle(sheetId, range, style => ({
      // General shows every digit, so decimal places only mean something once the cell is a number, as in Excel.
      numberFormat: style.numberFormat === 'general' ? 'number' : style.numberFormat,
      decimalPlaces: Math.max(0, Math.min(12, style.decimalPlaces + difference)),
    }))
  }

  /** Applies a style patch to every cell in the range; `undefined` values reset that property. */
  updateStyle(sheetId: number, range: CellRange, patch: Partial<CellStyle>): void {
    this.#updateStyle(sheetId, range, () => patch)
  }

  clearFormatting(sheetId: number, range: CellRange): void {
    this.#commit(() => this.#mapStyles(sheetId, range, () => defaultCellStyle))
  }

  /** Applies a per-cell patch. Changing the number format drops the file's own pattern, which no longer applies. */
  #updateStyle(sheetId: number, range: CellRange, patchFor: (style: CellStyle) => Partial<CellStyle>): void {
    this.#commit(() => this.#mapStyles(sheetId, range, style => {
      const patch = patchFor(style)
      const formatChanged = patch.numberFormat !== undefined || patch.decimalPlaces !== undefined
      return mergeStyle(style, formatChanged ? { ...patch, formatCode: undefined } : patch)
    }))
  }

  setCellInput(sheetId: number, [col, row]: Item, input: CellInput): void {
    if (this.getCellInput(sheetId, [col, row]) === String(input ?? '')) return
    this.#commit(() => this.#engine.setCellContents({ sheet: sheetId, col, row }, [[input]]))
  }

  /** Empties values and formulas in the range as one undo step, keeping formatting. */
  clearContents(sheetId: number, range: CellRange): void {
    const { width, height } = this.#engine.getSheetDimensions(sheetId)
    const right = Math.min(range.x + range.width, width)
    const bottom = Math.min(range.y + range.height, height)
    if (right <= range.x || bottom <= range.y) return
    this.#commit(() => this.#engine.setCellContents(
      { sheet: sheetId, col: range.x, row: range.y },
      Array.from({ length: bottom - range.y }, () => Array.from({ length: right - range.x }, () => null)),
    ))
  }

  /** Pastes tabular text and applies formats inferred from Excel or Google Sheets display values. */
  setBlock(sheetId: number, [startCol, startRow]: Item, values: readonly (readonly string[])[]): void {
    if (values.length === 0) return
    this.#commit(() => {
      const styles = this.#sheetStyles(this.#sheetName(sheetId))
      const parsed = values.map(row => row.map(parseCellText))
      this.#engine.setCellContents(
        { sheet: sheetId, col: startCol, row: startRow },
        parsed.map(row => row.map(cell => cell.input)),
      )
      parsed.forEach((row, rowOffset) => row.forEach((cell, colOffset) => {
        const key = styleKey(startCol + colOffset, startRow + rowOffset)
        if (cell.style === undefined) styles.delete(key)
        else styles.set(key, cell.style)
      }))
    })
  }

  /** Repeats source cells into a destination while HyperFormula translates relative formulas.
   * Copy/paste cannot run inside engine.batch (it forces evaluation), so each paste recomputes. */
  fillRange(sheetId: number, source: CellRange, destination: CellRange): void {
    this.#commit(() => {
      const styles = this.#sheetStyles(this.#sheetName(sheetId))
      this.#forEachCell(destination, (col, row) => {
        const sourceCol = source.x + ((col - source.x) % source.width + source.width) % source.width
        const sourceRow = source.y + ((row - source.y) % source.height + source.height) % source.height
        if (col === sourceCol && row === sourceRow) return
        this.#engine.copy({
          start: { sheet: sheetId, col: sourceCol, row: sourceRow },
          end: { sheet: sheetId, col: sourceCol, row: sourceRow },
        })
        this.#engine.paste({ sheet: sheetId, col, row })
        const sourceStyle = styles.get(styleKey(sourceCol, sourceRow))
        const targetKey = styleKey(col, row)
        if (sourceStyle === undefined) styles.delete(targetKey)
        else styles.set(targetKey, { ...sourceStyle })
      })
    })
  }

  setColumnWidth(sheetId: number, column: number, width: number): void {
    const layout = this.#layout(sheetId)
    const nextWidth = Math.max(48, Math.min(600, Math.round(width)))
    if (this.getColumnWidth(sheetId, column) === nextWidth) return
    this.#commit(() => layout.columnWidths.set(column, nextWidth))
  }

  insertRows(sheetId: number, index: number, count = 1): void {
    this.#commit(() => {
      const dimensions = this.#engine.getSheetDimensions(sheetId)
      if (index < dimensions.height) this.#engine.addRows(sheetId, [index, count])
      const layout = this.#layout(sheetId)
      layout.rowCount += count
      if (index < layout.freezeRows) {
        layout.freezeRows = Math.min(layout.rowCount, layout.freezeRows + count)
      }
      this.#shiftMetadata(this.#sheetName(sheetId), 'row', index, count, 0)
    })
  }

  deleteRows(sheetId: number, index: number, count = 1): void {
    const layout = this.#layout(sheetId)
    const removable = Math.min(count, layout.rowCount - 1, layout.rowCount - index)
    if (removable <= 0) return
    this.#commit(() => {
      const dimensions = this.#engine.getSheetDimensions(sheetId)
      const engineCount = Math.min(removable, Math.max(0, dimensions.height - index))
      if (engineCount > 0) this.#engine.removeRows(sheetId, [index, engineCount])
      if (index < layout.freezeRows) {
        layout.freezeRows -= Math.min(removable, layout.freezeRows - index)
      }
      layout.rowCount -= removable
      layout.freezeRows = Math.min(layout.freezeRows, layout.rowCount)
      this.#shiftMetadata(this.#sheetName(sheetId), 'row', index, -removable, removable)
    })
  }

  insertColumns(sheetId: number, index: number, count = 1): void {
    this.#commit(() => {
      const dimensions = this.#engine.getSheetDimensions(sheetId)
      if (index < dimensions.width) this.#engine.addColumns(sheetId, [index, count])
      const layout = this.#layout(sheetId)
      layout.columnCount += count
      if (index < layout.freezeColumns) {
        layout.freezeColumns = Math.min(layout.columnCount, layout.freezeColumns + count)
      }
      this.#shiftMetadata(this.#sheetName(sheetId), 'column', index, count, 0)
    })
  }

  deleteColumns(sheetId: number, index: number, count = 1): void {
    const layout = this.#layout(sheetId)
    const removable = Math.min(count, layout.columnCount - 1, layout.columnCount - index)
    if (removable <= 0) return
    this.#commit(() => {
      const dimensions = this.#engine.getSheetDimensions(sheetId)
      const engineCount = Math.min(removable, Math.max(0, dimensions.width - index))
      if (engineCount > 0) this.#engine.removeColumns(sheetId, [index, engineCount])
      if (index < layout.freezeColumns) {
        layout.freezeColumns -= Math.min(removable, layout.freezeColumns - index)
      }
      layout.columnCount -= removable
      layout.freezeColumns = Math.min(layout.freezeColumns, layout.columnCount)
      this.#shiftMetadata(this.#sheetName(sheetId), 'column', index, -removable, removable)
    })
  }

  hideRows(sheetId: number, index: number, count = 1): void {
    this.#hide(this.#layout(sheetId).hiddenRows, index, count)
  }

  hideColumns(sheetId: number, index: number, count = 1): void {
    this.#hide(this.#layout(sheetId).hiddenColumns, index, count)
  }

  /** Hides a run of rows or columns as one undo step; hiding what is already hidden records nothing. */
  #hide(hidden: Set<number>, index: number, count: number): void {
    const indexes = Array.from({ length: count }, (_, offset) => index + offset)
    if (indexes.every(value => hidden.has(value))) return
    this.#commit(() => indexes.forEach(value => hidden.add(value)))
  }

  showAllRows(sheetId: number): void {
    if (this.#layout(sheetId).hiddenRows.size === 0) return
    this.#commit(() => this.#layout(sheetId).hiddenRows.clear())
  }

  showAllColumns(sheetId: number): void {
    if (this.#layout(sheetId).hiddenColumns.size === 0) return
    this.#commit(() => this.#layout(sheetId).hiddenColumns.clear())
  }

  freezeColumns(sheetId: number, count: number): void {
    const layout = this.#layout(sheetId)
    const next = Math.max(0, Math.min(count, layout.columnCount))
    if (next === layout.freezeColumns) return
    this.#commit(() => { layout.freezeColumns = next })
  }

  freezeRows(sheetId: number, count: number): void {
    const layout = this.#layout(sheetId)
    const next = Math.max(0, Math.min(count, layout.rowCount))
    if (next === layout.freezeRows) return
    this.#commit(() => { layout.freezeRows = next })
  }

  addSheet(): WorkbookSheet {
    return this.#commit(() => {
      const name = this.#engine.addSheet()
      const id = this.#sheetId(name)
      this.#sheetOrder = [...this.#sheetOrder, name]
      this.#layouts.set(name, layoutFromData(blankLayout))
      return { id, name }
    })
  }

  renameSheet(sheetId: number, name: string): boolean {
    const trimmed = name.trim()
    if (!this.#engine.isItPossibleToRenameSheet(sheetId, trimmed)) return false
    this.#commit(() => {
      const previousName = this.#sheetName(sheetId)
      this.#engine.renameSheet(sheetId, trimmed)
      this.#sheetOrder = this.#sheetOrder.map(candidate => candidate === previousName ? trimmed : candidate)
      const layout = this.#layouts.get(previousName)
      if (layout !== undefined) {
        this.#layouts.delete(previousName)
        this.#layouts.set(trimmed, layout)
      }
      const styles = this.#styles.get(previousName)
      if (styles !== undefined) {
        this.#styles.delete(previousName)
        this.#styles.set(trimmed, styles)
      }
      this.#namedRanges.forEach((namedRange, key) => {
        if (namedRange.sheetName === previousName) {
          this.#namedRanges.set(key, { ...namedRange, sheetName: trimmed })
        }
      })
    })
    return true
  }

  removeSheet(sheetId: number): void {
    if (this.#sheetOrder.length === 1) throw new Error('A workbook must contain at least one sheet.')
    this.#commit(() => {
      const name = this.#sheetName(sheetId)
      const removedNames = [...this.#namedRanges.values()]
        .filter(namedRange => namedRange.sheetName === name)
        .map(namedRange => namedRange.name)
      removedNames.forEach(namedName => {
        this.#engine.removeNamedExpression(namedName)
        this.#namedRanges.delete(namedName.toLowerCase())
      })
      this.#engine.removeSheet(sheetId)
      this.#sheetOrder = this.#sheetOrder.filter(candidate => candidate !== name)
      this.#layouts.delete(name)
      this.#styles.delete(name)
    })
  }

  moveSheet(fromIndex: number, toIndex: number): void {
    if (fromIndex === toIndex) return
    const moved = this.#sheetOrder[fromIndex]
    if (moved === undefined) return
    this.#commit(() => {
      const next = [...this.#sheetOrder]
      next.splice(fromIndex, 1)
      next.splice(toIndex, 0, moved)
      this.#sheetOrder = next
    })
  }

  find(query: string, caseSensitive = false): readonly FindMatch[] {
    if (query === '') return []
    const needle = caseSensitive ? query : query.toLocaleLowerCase()
    return this.sheets().flatMap(sheet => this.#engine.getSheetSerialized(sheet.id).flatMap((row, rowIndex) =>
      row.flatMap((value, colIndex) => {
        const input = value === null || value === undefined ? '' : String(value)
        const candidate = caseSensitive ? input : input.toLocaleLowerCase()
        return candidate.includes(needle)
          ? [{ sheetId: sheet.id, sheetName: sheet.name, cell: [colIndex, rowIndex] as Item, input }]
          : []
      }),
    ))
  }

  replaceAll(query: string, replacement: string, caseSensitive = false): number {
    const matches = this.find(query, caseSensitive)
    if (matches.length === 0) return 0
    const pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), caseSensitive ? 'g' : 'gi')
    this.#commit(() => {
      this.#engine.batch(() => matches.forEach(match => {
        // A replacer function inserts the text literally; a replacement string would expand "$&" and "$1".
        const next = match.input.replace(pattern, () => replacement)
        this.#engine.setCellContents(
          { sheet: match.sheetId, col: match.cell[0], row: match.cell[1] },
          [[next]],
        )
      }))
    })
    return matches.length
  }

  namedRanges(): readonly NamedRange[] {
    return [...this.#namedRanges.values()].sort((left, right) => left.name.localeCompare(right.name))
  }

  addNamedRange(name: string, sheetId: number, range: CellRange): boolean {
    const normalized = name.trim()
    const namedRange = { name: normalized, sheetName: this.#sheetName(sheetId), range: cloneRange(range) }
    const formula = namedRangeFormula(namedRange)
    if (!this.#engine.isItPossibleToAddNamedExpression(normalized, formula)) return false
    this.#commit(() => {
      this.#engine.addNamedExpression(normalized, formula)
      this.#namedRanges.set(normalized.toLowerCase(), namedRange)
    })
    return true
  }

  removeNamedRange(name: string): void {
    const key = name.toLowerCase()
    if (!this.#namedRanges.has(key)) return
    this.#commit(() => {
      this.#engine.removeNamedExpression(name)
      this.#namedRanges.delete(key)
    })
  }

  /**
   * Reorders whole rows of the range by one column's values. Moving whole rows through the engine
   * keeps every formula's references pointing at the same data, including relative ones.
   */
  sortRows(sheetId: number, range: CellRange, byColumn: number, ascending: boolean): void {
    const height = this.#engine.getSheetDimensions(sheetId).height
    const start = range.y
    const end = Math.min(range.y + range.height, height)
    if (end - start < 2) return
    const rows = Array.from({ length: end - start }, (_, index) => start + index)
    const values = new Map(rows.map(row => [row, this.#engine.getCellValue({ sheet: sheetId, col: byColumn, row })]))
    const sorted = [...rows].sort((left, right) => compareForSort(values.get(left) ?? null, values.get(right) ?? null, ascending))
    if (sorted.every((row, index) => row === rows[index])) return
    // newOrder[current row] = row's new position.
    const newOrder = Array.from({ length: height }, (_, index) => index)
    sorted.forEach((row, index) => { newOrder[row] = start + index })
    this.#commit(() => {
      this.#engine.setRowOrder(sheetId, newOrder)
      const name = this.#sheetName(sheetId)
      const styles = this.#sheetStyles(name)
      const moved = new Map<string, CellStyle>()
      styles.forEach((style, key) => {
        const [col, row] = parseStyleKey(key)
        moved.set(styleKey(col, row >= start && row < end ? newOrder[row] ?? row : row), style)
      })
      this.#styles.set(name, moved)
    })
  }

  /** The block of rows below a header that a single-cell sort should reorder. */
  sortRegion(sheetId: number): CellRange | undefined {
    const { width, height } = this.#engine.getSheetDimensions(sheetId)
    if (height < 2) return undefined
    const firstRow = Array.from({ length: width }, (_, col) => this.#engine.getCellValue({ sheet: sheetId, col, row: 0 }))
    const headerIsText = firstRow.some(value => typeof value === 'string' && value !== '')
      && firstRow.every(value => value === null || typeof value === 'string')
    const headerIsBold = firstRow.some((_, col) => this.#style(sheetId, col, 0).bold === true)
    const start = headerIsText || headerIsBold ? 1 : 0
    return height - start < 2 ? undefined : { x: 0, y: start, width, height: height - start }
  }

  selectionStats(sheetId: number, range: CellRange): SelectionStats {
    const { width, height } = this.#engine.getSheetDimensions(sheetId)
    let count = 0
    let numericCount = 0
    let sum = 0
    let min = Number.POSITIVE_INFINITY
    let max = Number.NEGATIVE_INFINITY
    for (let row = range.y; row < Math.min(range.y + range.height, height); row += 1) {
      for (let col = range.x; col < Math.min(range.x + range.width, width); col += 1) {
        const value = this.#engine.getCellValue({ sheet: sheetId, col, row })
        if (value === null || value === '') continue
        count += 1
        if (typeof value !== 'number') continue
        numericCount += 1
        sum += value
        min = Math.min(min, value)
        max = Math.max(max, value)
      }
    }
    return { count, numericCount, sum, min, max }
  }

  /** Writes =SUM() of the unbroken run of numbers directly above the cell, else directly left. False if none. */
  autoSum(sheetId: number, [col, row]: Item): boolean {
    const isNumber = (c: number, r: number): boolean =>
      typeof this.#engine.getCellValue({ sheet: sheetId, col: c, row: r }) === 'number'
    let top = row
    while (top > 0 && isNumber(col, top - 1)) top -= 1
    if (top < row) {
      this.setCellInput(sheetId, [col, row], `=SUM(${columnName(col)}${top + 1}:${columnName(col)}${row})`)
      return true
    }
    let first = col
    while (first > 0 && isNumber(first - 1, row)) first -= 1
    if (first < col) {
      this.setCellInput(sheetId, [col, row], `=SUM(${columnName(first)}${row + 1}:${columnName(col - 1)}${row + 1})`)
      return true
    }
    return false
  }

  /** The used area of a sheet, for exports. */
  usedSize(sheetId: number): { readonly width: number; readonly height: number } {
    return this.#engine.getSheetDimensions(sheetId)
  }

  /** The workbook as plain data, with computed values, for saving to a file. */
  toData(): WorkbookData {
    return this.#serialize(true)
  }

  isDirty(): boolean {
    return this.#depth !== this.#savedDepth
  }

  /** The undo position a save is about to write; hand it to `markSaved` once the write finishes. */
  savePoint(): number {
    return this.#depth
  }

  /**
   * Records `point` as saved, so undoing back to it clears the edited flag. Taking the point before
   * the write keeps an edit made during the write marked as unsaved.
   */
  markSaved(point = this.#depth): void {
    this.#savedDepth = point
    this.#emit()
  }

  undo(): void {
    const previous = this.#undoStack.pop()
    if (previous === undefined) return
    this.#redoStack.push(this.#serialize(false))
    this.#load(previous)
    this.#depth -= 1
    this.#emit()
  }

  redo(): void {
    const next = this.#redoStack.pop()
    if (next === undefined) return
    this.#undoStack.push(this.#serialize(false))
    this.#load(next)
    this.#depth += 1
    this.#emit()
  }

  canUndo(): boolean {
    return this.#undoStack.length > 0
  }

  canRedo(): boolean {
    return this.#redoStack.length > 0
  }

  functionSuggestions(prefix: string): readonly FunctionListEntry[] {
    const normalized = prefix.toUpperCase()
    return this.#engine.getAvailableFunctions()
      .filter(entry => entry.localizedName.startsWith(normalized))
      .slice(0, 8)
  }

  functionDetails(canonicalName: string): FunctionDetails | undefined {
    return this.#engine.getFunctionDetails(canonicalName)
  }

  #style(sheetId: number, col: number, row: number): CellStyle {
    return this.#styles.get(this.#sheetName(sheetId))?.get(styleKey(col, row)) ?? defaultCellStyle
  }

  #layout(sheetId: number): SheetLayout {
    const name = this.#sheetName(sheetId)
    const layout = this.#layouts.get(name)
    if (layout === undefined) throw new Error(`Workbook layout is missing: ${name}`)
    return layout
  }

  #sheetStyles(name: string): Map<string, CellStyle> {
    const current = this.#styles.get(name)
    if (current !== undefined) return current
    const styles = new Map<string, CellStyle>()
    this.#styles.set(name, styles)
    return styles
  }

  #sheetId(name: string): number {
    const id = this.#engine.getSheetId(name)
    if (id === undefined) throw new Error(`Workbook sheet is missing: ${name}`)
    return id
  }

  #sheetName(id: number): string {
    const name = this.#engine.getSheetName(id)
    if (name === undefined) throw new Error(`Workbook sheet is missing: ${id}`)
    return name
  }

  #forEachCell(range: CellRange, callback: (col: number, row: number) => void): void {
    for (let row = range.y; row < range.y + range.height; row += 1) {
      for (let col = range.x; col < range.x + range.width; col += 1) callback(col, row)
    }
  }

  #mapStyles(sheetId: number, range: CellRange, map: (style: CellStyle) => CellStyle): void {
    const styles = this.#sheetStyles(this.#sheetName(sheetId))
    this.#forEachCell(range, (col, row) => {
      const key = styleKey(col, row)
      const next = map(styles.get(key) ?? defaultCellStyle)
      if (isDefaultStyle(next)) styles.delete(key)
      else styles.set(key, next)
    })
  }

  #shiftMetadata(
    sheetName: string,
    axis: 'row' | 'column',
    index: number,
    offset: number,
    removedCount: number,
  ): void {
    const removedRangeKeys: string[] = []
    const styles = this.#sheetStyles(sheetName)
    const shiftedStyles = new Map<string, CellStyle>()
    styles.forEach((style, key) => {
      const [col, row] = parseStyleKey(key)
      const coordinate = axis === 'row' ? row : col
      if (removedCount > 0 && coordinate >= index && coordinate < index + removedCount) return
      const shifted = coordinate >= index + removedCount ? coordinate + offset : coordinate
      shiftedStyles.set(
        styleKey(axis === 'column' ? shifted : col, axis === 'row' ? shifted : row),
        style,
      )
    })
    this.#styles.set(sheetName, shiftedStyles)

    const layout = this.#layouts.get(sheetName)
    if (layout === undefined) return
    const hidden = axis === 'row' ? layout.hiddenRows : layout.hiddenColumns
    const shiftedHidden = [...hidden].flatMap(coordinate => {
      if (removedCount > 0 && coordinate >= index && coordinate < index + removedCount) return []
      return [coordinate >= index + removedCount ? coordinate + offset : coordinate]
    })
    hidden.clear()
    shiftedHidden.forEach(coordinate => hidden.add(coordinate))

    if (axis === 'column') {
      const shiftedWidths = new Map<number, number>()
      layout.columnWidths.forEach((width, coordinate) => {
        if (removedCount > 0 && coordinate >= index && coordinate < index + removedCount) return
        shiftedWidths.set(coordinate >= index + removedCount ? coordinate + offset : coordinate, width)
      })
      layout.columnWidths.clear()
      shiftedWidths.forEach((width, coordinate) => layout.columnWidths.set(coordinate, width))
    }

    this.#namedRanges.forEach((namedRange, key) => {
      if (namedRange.sheetName !== sheetName) return
      const range = cloneRange(namedRange.range)
      const start = axis === 'row' ? range.y : range.x
      const size = axis === 'row' ? range.height : range.width
      if (removedCount === 0) {
        if (index <= start) {
          if (axis === 'row') range.y += offset
          else range.x += offset
        } else if (index < start + size) {
          if (axis === 'row') range.height += offset
          else range.width += offset
        }
      } else {
        const end = start + size
        const removedEnd = index + removedCount
        if (start >= index && end <= removedEnd) {
          removedRangeKeys.push(key)
          return
        }
        const overlap = Math.max(0, Math.min(end, removedEnd) - Math.max(start, index))
        if (removedEnd <= start) {
          if (axis === 'row') range.y += offset
          else range.x += offset
        }
        if (axis === 'row') range.height = Math.max(1, range.height - overlap)
        else range.width = Math.max(1, range.width - overlap)
      }
      this.#namedRanges.set(key, { ...namedRange, range })
    })
    // A range whose every cell was deleted no longer names anything.
    removedRangeKeys.forEach(key => {
      const namedRange = this.#namedRanges.get(key)
      if (namedRange !== undefined) this.#engine.removeNamedExpression(namedRange.name)
      this.#namedRanges.delete(key)
    })
  }

  #serialize(withValues: boolean): WorkbookData {
    return {
      sheets: this.#sheetOrder.map((name): SheetData => {
        const id = this.#sheetId(name)
        const layout = this.#layout(id)
        return {
          name,
          cells: this.#engine.getSheetSerialized(id).map(row => row.map(serializedInput)),
          values: withValues ? this.#engine.getSheetValues(id).map(row => row.map(plainValue)) : undefined,
          styles: [...(this.#styles.get(name)?.entries() ?? [])].map(([key, style]) => {
            const [col, row] = parseStyleKey(key)
            return [col, row, style] as const
          }),
          layout: {
            rowCount: layout.rowCount,
            columnCount: layout.columnCount,
            columnWidths: [...layout.columnWidths.entries()],
            hiddenRows: [...layout.hiddenRows],
            hiddenColumns: [...layout.hiddenColumns],
            freezeRows: layout.freezeRows,
            freezeColumns: layout.freezeColumns,
          },
        }
      }),
      namedRanges: [...this.#namedRanges.values()].map(namedRange => ({ ...namedRange, range: cloneRange(namedRange.range) })),
    }
  }

  /** Replaces all state with the given data; used by construction, undo and redo. */
  #load(data: WorkbookData): void {
    const engine = HyperFormula.buildEmpty(workbookConfig)
    engine.batch(() => {
      data.sheets.forEach(sheet => engine.addSheet(sheet.name))
      data.namedRanges.forEach(namedRange => {
        const formula = namedRangeFormula(namedRange)
        if (engine.isItPossibleToAddNamedExpression(namedRange.name, formula)) engine.addNamedExpression(namedRange.name, formula)
      })
      data.sheets.forEach(sheet => {
        engine.setSheetContent(this.#requiredSheetId(engine, sheet.name), sheet.cells.map(row => [...row]))
      })
    })
    this.#engine = engine
    this.#sheetOrder = data.sheets.map(sheet => sheet.name)
    this.#layouts.clear()
    this.#styles.clear()
    data.sheets.forEach(sheet => {
      this.#layouts.set(sheet.name, layoutFromData(sheet.layout))
      this.#styles.set(sheet.name, new Map(sheet.styles.map(([col, row, style]) => [styleKey(col, row), style])))
    })
    this.#namedRanges.clear()
    data.namedRanges
      .filter(namedRange => engine.getNamedExpression(namedRange.name) !== undefined)
      .forEach(namedRange => this.#namedRanges.set(namedRange.name.toLowerCase(), { ...namedRange, range: cloneRange(namedRange.range) }))
  }

  #requiredSheetId(engine: HyperFormula, name: string): number {
    const id = engine.getSheetId(name)
    if (id === undefined) throw new Error(`Workbook sheet is missing while loading: ${name}`)
    return id
  }

  #commit<Result>(mutation: () => Result): Result {
    const before = this.#serialize(false)
    const result = mutation()
    this.#undoStack.push(before)
    if (this.#undoStack.length > historyLimit) this.#undoStack.shift()
    this.#redoStack.length = 0
    // Branching off an undone state makes the saved state (on the discarded branch) unreachable.
    if (this.#savedDepth !== undefined && this.#savedDepth > this.#depth) this.#savedDepth = undefined
    this.#depth += 1
    this.#emit()
    return result
  }

  #emit(): void {
    this.#revision += 1
    this.#listeners.forEach(listener => listener())
  }
}
