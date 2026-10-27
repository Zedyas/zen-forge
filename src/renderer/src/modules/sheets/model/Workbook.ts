import type { Item } from '@glideapps/glide-data-grid'
import {
  HyperFormula,
  type CellValue,
  type FunctionDetails,
  type FunctionListEntry,
  type RawCellContent,
} from 'hyperformula'
import { columnName, sheetNameProblem, sheetRangeReference } from './address'
import { parseCellText, shownInput } from './cell-parse'
import { fillSeries } from './fill-series'
import { formatCellValue, isDateStyle } from './format'
import { shiftFormula } from './references'
import {
  defaultCellStyle,
  defaultColumnCount,
  defaultRowCount,
  defaultRowHeight,
  isDefaultStyle,
  type CellInput,
  type CellRange,
  type CellStyle,
  type CopiedCells,
  type NamedRange,
  type NumberFormat,
  type PastedCell,
  type SheetData,
  type SheetLayoutData,
  type WorkbookData,
} from './workbook-data'

export type { CellInput, CellRange, CellStyle, CopiedCells, NamedRange, NumberFormat, PastedCell, WorkbookData } from './workbook-data'

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

/** One column (fill down or up) or row (fill right or left) of a fill-handle drag, as workbook cells. */
export interface FillLine {
  /** The selected cells the drag started from, top to bottom or left to right. */
  readonly seeds: readonly Item[]
  /** The cells the drag added, nearest the seeds first; hidden rows and columns are left out. */
  readonly targets: readonly Item[]
}

interface SheetLayout {
  rowCount: number
  columnCount: number
  readonly columnWidths: Map<number, number>
  readonly rowHeights: Map<number, number>
  readonly hiddenRows: Set<number>
  readonly hiddenColumns: Set<number>
  freezeRows: number
  freezeColumns: number
  merges: readonly CellRange[]
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
  // Excel's grid; the engine's own default stops at 40,000 rows, so a larger file would not open.
  maxRows: 1_048_576,
  maxColumns: 16_384,
  // Undo history is owned by the snapshot stacks below; the engine keeps none.
  undoLimit: 0,
} as const

const defaultColumnWidth = 120
const historyLimit = 100

const blankLayout: SheetLayoutData = {
  rowCount: defaultRowCount,
  columnCount: defaultColumnCount,
  columnWidths: [],
  rowHeights: [],
  hiddenRows: [],
  hiddenColumns: [],
  freezeRows: 0,
  freezeColumns: 0,
  merges: [],
}

/**
 * The height a cell's text needs at a column width, in pixels. Measuring needs the screen's fonts, so
 * the grid supplies it; the workbook uses it to fit rows to wrapped and larger text.
 */
export type CellMeasure = (text: string, style: CellStyle, width: number) => number

/** Style properties that change how tall a cell's text is. */
const textLayoutKeys: readonly (keyof CellStyle)[] = ['wrap', 'fontSize', 'fontFamily']

function layoutFromData(layout: SheetLayoutData): SheetLayout {
  return {
    rowCount: Math.max(layout.rowCount, defaultRowCount),
    columnCount: Math.max(layout.columnCount, defaultColumnCount),
    columnWidths: new Map(layout.columnWidths),
    rowHeights: new Map(layout.rowHeights),
    hiddenRows: new Set(layout.hiddenRows),
    hiddenColumns: new Set(layout.hiddenColumns),
    freezeRows: layout.freezeRows,
    freezeColumns: layout.freezeColumns,
    merges: layout.merges.map(cloneRange),
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

/** A raw input as the formula bar shows it and Find searches it, without the mark that keeps every text cell text. */
function shownCellInput(serialized: RawCellContent): string {
  if (serialized instanceof Date) return serialized.toISOString()
  if (typeof serialized === 'string') return shownInput(serialized)
  return serialized === null || serialized === undefined ? '' : String(serialized)
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

function intersects(left: CellRange, right: CellRange): boolean {
  return left.x < right.x + right.width && right.x < left.x + left.width
    && left.y < right.y + right.height && right.y < left.y + left.height
}

function containsRange(outer: CellRange, inner: CellRange): boolean {
  return inner.x >= outer.x && inner.y >= outer.y
    && inner.x + inner.width <= outer.x + outer.width && inner.y + inner.height <= outer.y + outer.height
}

function containsCell(range: CellRange, [col, row]: Item): boolean {
  return col >= range.x && col < range.x + range.width && row >= range.y && row < range.y + range.height
}

/**
 * A range after rows or columns are inserted (`removedCount` 0, `offset` the count) or deleted
 * (`offset` = -`removedCount`) at `index`: it moves, grows when the insert falls inside it, shrinks
 * when part of it is deleted, and is gone (undefined) when all of it is.
 */
function shiftRange(range: CellRange, axis: 'row' | 'column', index: number, offset: number, removedCount: number): CellRange | undefined {
  const start = axis === 'row' ? range.y : range.x
  const size = axis === 'row' ? range.height : range.width
  let nextStart = start
  let nextSize = size
  if (removedCount === 0) {
    if (index <= start) nextStart += offset
    else if (index < start + size) nextSize += offset
  } else {
    const removedEnd = index + removedCount
    const overlap = Math.max(0, Math.min(start + size, removedEnd) - Math.max(start, index))
    if (overlap === size) return undefined
    nextStart = start >= removedEnd ? start - removedCount : Math.min(start, index)
    nextSize = size - overlap
  }
  return axis === 'row'
    ? { x: range.x, y: nextStart, width: range.width, height: nextSize }
    : { x: nextStart, y: range.y, width: nextSize, height: range.height }
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
    return shownCellInput(this.#engine.getCellSerialized({ sheet: sheetId, col, row }))
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

  getRowHeight(sheetId: number, row: number): number {
    return this.#layout(sheetId).rowHeights.get(row) ?? defaultRowHeight
  }

  /** Sets rows to `height` pixels as one undo step; undefined puts them back to the default height. */
  setRowHeight(sheetId: number, rows: readonly number[], height: number | undefined): void {
    const heights = this.#layout(sheetId).rowHeights
    const next = height === undefined ? undefined : Math.max(12, Math.min(600, Math.round(height)))
    if (rows.every(row => heights.get(row) === next)) return
    this.#commit(() => rows.forEach(row => {
      if (next === undefined || next === defaultRowHeight) heights.delete(row)
      else heights.set(row, next)
    }))
  }

  /** Makes each row just tall enough for its tallest cell, as double-clicking a row border does in Excel. */
  fitRows(sheetId: number, rows: readonly number[], measure: CellMeasure): void {
    this.#commit(() => this.#fitRows(sheetId, rows, measure, false))
  }

  /** `growOnly` keeps a row that is already tall enough, such as one sized by hand, as it is. */
  #fitRows(sheetId: number, rows: readonly number[], measure: CellMeasure, growOnly: boolean): void {
    const layout = this.#layout(sheetId)
    const styles = this.#styles.get(this.#sheetName(sheetId))
    // A loop, not Math.max(...styles): a spread of more than about 125,000 styled cells overflows the call stack.
    let columns = this.#engine.getSheetDimensions(sheetId).width
    styles?.forEach((_, key) => { columns = Math.max(columns, parseStyleKey(key)[0] + 1) })
    new Set(rows).forEach(row => {
      let needed = defaultRowHeight
      for (let col = 0; col < columns; col += 1) {
        const style = this.#style(sheetId, col, row)
        const text = formatCellValue(plainValue(this.#engine.getCellValue({ sheet: sheetId, col, row })), style)
        if (text !== '') needed = Math.max(needed, measure(text, style, this.getColumnWidth(sheetId, col)))
      }
      const height = Math.min(600, Math.ceil(needed))
      if (growOnly && height <= (layout.rowHeights.get(row) ?? defaultRowHeight)) return
      if (height === defaultRowHeight) layout.rowHeights.delete(row)
      else layout.rowHeights.set(row, height)
    })
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

  setNumberFormat(sheetId: number, areas: readonly CellRange[], format: NumberFormat): void {
    this.updateStyle(sheetId, areas, { numberFormat: format })
  }

  changeDecimalPlaces(sheetId: number, areas: readonly CellRange[], difference: -1 | 1): void {
    this.#updateStyle(sheetId, areas, style => ({
      // General shows every digit, so decimal places only mean something once the cell is a number, as in Excel.
      numberFormat: style.numberFormat === 'general' ? 'number' : style.numberFormat,
      decimalPlaces: Math.max(0, Math.min(12, style.decimalPlaces + difference)),
    }))
  }

  /**
   * Applies a style patch to every cell of the areas; `undefined` values reset that property. Areas are
   * the rectangles of one selection, which leave out hidden rows and columns. With `fit`, a change to
   * wrapping or the font refits the rows to their text in the same undo step.
   */
  updateStyle(sheetId: number, areas: readonly CellRange[], patch: Partial<CellStyle>, fit?: CellMeasure): void {
    this.#updateStyle(sheetId, areas, () => patch)
    if (fit === undefined || !textLayoutKeys.some(key => key in patch)) return
    const rows = areas.flatMap(area => Array.from({ length: area.height }, (_, offset) => area.y + offset))
    this.#amend(() => this.#fitRows(sheetId, rows, fit, false))
  }

  /** Notes and links are content, as in Excel, so they stay. */
  clearFormatting(sheetId: number, areas: readonly CellRange[]): void {
    this.#commit(() => this.#mapStyles(sheetId, areas, style => mergeStyle(defaultCellStyle, { note: style.note, link: style.link })))
  }

  /**
   * Applies a per-cell patch. Changing the number format drops the cell's format code, which no longer
   * applies, unless the patch sets a new one.
   */
  #updateStyle(sheetId: number, areas: readonly CellRange[], patchFor: (style: CellStyle) => Partial<CellStyle>): void {
    this.#commit(() => this.#mapStyles(sheetId, areas, style => {
      const patch = patchFor(style)
      const formatChanged = (patch.numberFormat !== undefined || patch.decimalPlaces !== undefined) && !('formatCode' in patch)
      return mergeStyle(style, formatChanged ? { ...patch, formatCode: undefined } : patch)
    }))
  }

  /**
   * With `fit`, a row too short for the new text of a wrapped or larger-font cell grows to show it.
   * False when a merged cell covers the cell, as only a merged cell's upper-left cell holds a value.
   */
  setCellInput(sheetId: number, [col, row]: Item, input: CellInput, fit?: CellMeasure): boolean {
    if (this.isCovered(sheetId, [col, row])) return false
    const current = this.#engine.getCellSerialized({ sheet: sheetId, col, row })
    if (current === input || shownCellInput(current) === String(input ?? '')) return true
    this.#commit(() => {
      this.#engine.setCellContents({ sheet: sheetId, col, row }, [[input]])
      const style = this.#style(sheetId, col, row)
      if (fit !== undefined && textLayoutKeys.some(key => style[key] !== undefined)) this.#fitRows(sheetId, [row], fit, true)
    })
    return true
  }

  /** Empties values and formulas in the areas as one undo step, keeping formatting. */
  clearContents(sheetId: number, areas: readonly CellRange[]): void {
    const filled = this.#filledAreas(sheetId, areas)
    if (filled.length > 0) this.#commit(() => this.#clearValues(sheetId, filled))
  }

  /** Pastes tabular text and applies formats inferred from Excel or Google Sheets display values. */
  setBlock(sheetId: number, [startCol, startRow]: Item, values: readonly (readonly string[])[]): void {
    this.pasteCells(sheetId, values.flatMap((row, rowOffset) => row.map((text, colOffset): PastedCell => ({
      target: [startCol + colOffset, startRow + rowOffset],
      ...parseCellText(text),
    }))))
  }

  /** The raw inputs and styles of the given columns and rows, for the clipboard. */
  copyCells(sheetId: number, columns: readonly number[], rows: readonly number[]): CopiedCells {
    return {
      columns,
      rows,
      cells: rows.map(row => columns.map(col => {
        const style = this.#styles.get(this.#sheetName(sheetId))?.get(styleKey(col, row))
        return { input: serializedInput(this.#engine.getCellSerialized({ sheet: sheetId, col, row })), ...(style === undefined ? {} : { style }) }
      })),
    }
  }

  /**
   * Writes pasted cells and their styles as one undo step; copied formulas keep pointing the same distance away.
   * With `cutSheetId`, the cells were cut from that sheet, so their sources are emptied first, styles included.
   * Cells a merged cell covers are left out: only its upper-left cell holds a value.
   */
  pasteCells(sheetId: number, cells: readonly PastedCell[], cutSheetId?: number): void {
    if (cells.length === 0) return
    this.#commit(() => {
      if (cutSheetId !== undefined) {
        this.#clearValues(cutSheetId, this.#filledAreas(cutSheetId, cells.flatMap(({ source }) =>
          source === undefined ? [] : [{ x: source[0], y: source[1], width: 1, height: 1 }])))
        const cutStyles = this.#sheetStyles(this.#sheetName(cutSheetId))
        cells.forEach(({ source }) => { if (source !== undefined) cutStyles.delete(styleKey(source[0], source[1])) })
      }
      const styles = this.#sheetStyles(this.#sheetName(sheetId))
      const written = cells.filter(({ target }) => !this.isCovered(sheetId, target))
      this.#engine.batch(() => written.forEach(({ target: [col, row], input, style, source }) => {
        const moved = typeof input === 'string' && source !== undefined ? shiftFormula(input, col - source[0], row - source[1]) : input
        this.#engine.setCellContents({ sheet: sheetId, col, row }, [[moved]])
        if (style === undefined || isDefaultStyle(style)) styles.delete(styleKey(col, row))
        else styles.set(styleKey(col, row), style)
      }))
    })
  }

  /**
   * Moves cells, as Cut and then Paste do: formulas anywhere that referred to them follow them, and their
   * styles go along, as one undo step. False when the engine cannot, such as when it would split an array.
   */
  moveCells(fromSheetId: number, range: CellRange, toSheetId: number, [col, row]: Item): boolean {
    const source = {
      start: { sheet: fromSheetId, col: range.x, row: range.y },
      end: { sheet: fromSheetId, col: range.x + range.width - 1, row: range.y + range.height - 1 },
    }
    const destination = { sheet: toSheetId, col, row }
    if (!this.#engine.isItPossibleToMoveCells(source, destination)) return false
    this.#commit(() => {
      this.#engine.moveCells(source, destination)
      const from = this.#sheetStyles(this.#sheetName(fromSheetId))
      const to = this.#sheetStyles(this.#sheetName(toSheetId))
      const moved: [string, CellStyle | undefined][] = []
      this.#forEachCell(range, (sourceCol, sourceRow) => {
        const key = styleKey(sourceCol, sourceRow)
        moved.push([styleKey(col + sourceCol - range.x, row + sourceRow - range.y), from.get(key)])
        from.delete(key)
      })
      moved.forEach(([key, style]) => {
        if (style === undefined) to.delete(key)
        else to.set(key, style)
      })
      // Blocks inside the moved cells go with them; blocks they land on are unmerged.
      const fromLayout = this.#layout(fromSheetId)
      const toLayout = this.#layout(toSheetId)
      const target = { x: col, y: row, width: range.width, height: range.height }
      const carried = fromLayout.merges.filter(merge => containsRange(range, merge))
      fromLayout.merges = fromLayout.merges.filter(merge => !carried.includes(merge))
      toLayout.merges = [
        ...toLayout.merges.filter(merge => !intersects(merge, target)),
        ...carried.map(merge => ({ ...merge, x: merge.x + col - range.x, y: merge.y + row - range.y })),
      ]
      // The engine moves a name whose whole range was cut; its saved definition moves the same way.
      const fromName = this.#sheetName(fromSheetId)
      this.#namedRanges.forEach((namedRange, key) => {
        const { x, y, width, height } = namedRange.range
        if (namedRange.sheetName !== fromName || !containsRange(range, namedRange.range)) return
        const moved = { x: x + col - range.x, y: y + row - range.y, width, height }
        this.#namedRanges.set(key, { ...namedRange, sheetName: this.#sheetName(toSheetId), range: moved })
      })
    })
    return true
  }

  /** Empties values, formulas and formatting in the areas as one undo step. */
  clearCells(sheetId: number, areas: readonly CellRange[]): void {
    this.#commit(() => {
      this.#clearValues(sheetId, this.#filledAreas(sheetId, areas))
      this.#mapStyles(sheetId, areas, () => defaultCellStyle)
    })
  }

  /** The parts of the areas inside the sheet's used area: the only cells that can hold values. */
  #filledAreas(sheetId: number, areas: readonly CellRange[]): readonly CellRange[] {
    const { width, height } = this.#engine.getSheetDimensions(sheetId)
    return areas.flatMap(area => {
      const right = Math.min(area.x + area.width, width)
      const bottom = Math.min(area.y + area.height, height)
      return right <= area.x || bottom <= area.y ? [] : [{ x: area.x, y: area.y, width: right - area.x, height: bottom - area.y }]
    })
  }

  #clearValues(sheetId: number, areas: readonly CellRange[]): void {
    this.#engine.batch(() => areas.forEach(area => this.#engine.setCellContents(
      { sheet: sheetId, col: area.x, row: area.y },
      Array.from({ length: area.height }, () => Array.from({ length: area.width }, () => null)),
    )))
  }

  /**
   * The fill handle, as one undo step: each line continues its seeds as a series (1, 2 → 3, 4; Jan →
   * Feb; Item 1 → Item 2) or repeats them, with copied formulas' references moved and styles copied.
   * `backward` fills up or left, so the targets count down from the seeds. False, writing nothing, when
   * a target is a cell a merged cell covers.
   */
  fillCells(sheetId: number, lines: readonly FillLine[], backward: boolean): boolean {
    if (lines.some(({ targets }) => targets.some(target => this.isCovered(sheetId, target)))) return false
    // The engine cannot be read inside a batch, so every seed is read before anything is written.
    const writes = lines.flatMap(({ seeds, targets }) => {
      const filled = fillSeries(
        seeds.map(([col, row]) => ({
          input: serializedInput(this.#engine.getCellSerialized({ sheet: sheetId, col, row })),
          value: plainValue(this.#engine.getCellValue({ sheet: sheetId, col, row })),
          date: isDateStyle(this.#style(sheetId, col, row)),
        })),
        targets.map((_, index) => backward ? -1 - index : seeds.length + index),
      )
      return targets.flatMap((target, index) => {
        const cell = filled[index]
        const source = cell === undefined ? undefined : seeds[cell.seed]
        return cell === undefined || source === undefined ? [] : [{ target, source, input: cell.input }]
      })
    })
    if (writes.length === 0) return true
    this.#commit(() => {
      const styles = this.#sheetStyles(this.#sheetName(sheetId))
      this.#engine.batch(() => writes.forEach(({ target: [col, row], source, input }) => {
        const moved = typeof input === 'string' ? shiftFormula(input, col - source[0], row - source[1]) : input
        this.#engine.setCellContents({ sheet: sheetId, col, row }, [[moved]])
        const style = styles.get(styleKey(source[0], source[1]))
        if (style === undefined) styles.delete(styleKey(col, row))
        else styles.set(styleKey(col, row), style)
      }))
    })
    return true
  }

  merges(sheetId: number): readonly CellRange[] {
    return this.#layout(sheetId).merges
  }

  /** Whether a merged cell covers the cell: it is in a merged block but not the block's upper-left cell. */
  isCovered(sheetId: number, cell: Item): boolean {
    const merge = this.mergeAt(sheetId, cell)
    return merge !== undefined && (merge.x !== cell[0] || merge.y !== cell[1])
  }

  /** The merged block holding the cell, if any. */
  mergeAt(sheetId: number, cell: Item): CellRange | undefined {
    return this.#layout(sheetId).merges.find(merge => containsCell(merge, cell))
  }

  /** The blocks a merge of `range` makes: the whole range, or with `across`, one per row. One-cell blocks are left out. */
  #mergeBlocks(range: CellRange, across: boolean): readonly CellRange[] {
    const blocks = across
      ? Array.from({ length: range.height }, (_, offset) => ({ x: range.x, y: range.y + offset, width: range.width, height: 1 }))
      : [cloneRange(range)]
    return blocks.filter(block => block.width * block.height > 1)
  }

  /** How many values a merge would discard: Excel keeps only each block's top-left value. */
  mergeDiscards(sheetId: number, range: CellRange, across: boolean): number {
    return this.#mergeBlocks(range, across).reduce((total, block) => {
      let count = 0
      this.#forEachCell(block, (col, row) => {
        const value = this.#engine.getCellValue({ sheet: sheetId, col, row })
        if ((col !== block.x || row !== block.y) && value !== null && value !== '') count += 1
      })
      return total + count
    }, 0)
  }

  /**
   * Merges the range, or with `across` each of its rows, as one undo step. Blocks it overlaps are
   * replaced, and every value but each block's top-left is discarded, as in Excel.
   */
  mergeCells(sheetId: number, range: CellRange, across: boolean): void {
    const blocks = this.#mergeBlocks(range, across)
    if (blocks.length === 0) return
    this.#commit(() => {
      const layout = this.#layout(sheetId)
      const { width, height } = this.#engine.getSheetDimensions(sheetId)
      this.#engine.batch(() => blocks.forEach(block => this.#forEachCell(block, (col, row) => {
        if ((col !== block.x || row !== block.y) && col < width && row < height) this.#engine.setCellContents({ sheet: sheetId, col, row }, [[null]])
      })))
      layout.merges = [...layout.merges.filter(merge => !intersects(merge, range)), ...blocks]
    })
  }

  unmergeCells(sheetId: number, range: CellRange): void {
    const layout = this.#layout(sheetId)
    if (!layout.merges.some(merge => intersects(merge, range))) return
    this.#commit(() => { layout.merges = layout.merges.filter(merge => !intersects(merge, range)) })
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

  /** Named Sheet and the next number no sheet uses; the engine's own choice can be taken (Sheet1 and Sheet3 make Sheet3). */
  addSheet(): WorkbookSheet {
    const taken = new Set(this.#sheetOrder.map(name => name.toLowerCase()))
    let number = this.#sheetOrder.length + 1
    while (taken.has(`sheet${number}`)) number += 1
    return this.#commit(() => {
      const name = this.#engine.addSheet(`Sheet${number}`)
      const id = this.#sheetId(name)
      this.#sheetOrder = [...this.#sheetOrder, name]
      this.#layouts.set(name, layoutFromData(blankLayout))
      return { id, name }
    })
  }

  /** False when the name breaks Excel's rules (see `sheetNameProblem`) or is unchanged. */
  renameSheet(sheetId: number, name: string): boolean {
    const trimmed = name.trim()
    const previousName = this.#sheetName(sheetId)
    if (trimmed === previousName) return false
    if (sheetNameProblem(trimmed, this.#sheetOrder.filter(candidate => candidate !== previousName)) !== undefined) return false
    if (!this.#engine.isItPossibleToRenameSheet(sheetId, trimmed)) return false
    this.#commit(() => {
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
        const input = shownCellInput(value)
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
  sortRows(sheetId: number, range: CellRange, byColumn: number, ascending: boolean): 'sorted' | 'merged' {
    const height = this.#engine.getSheetDimensions(sheetId).height
    const start = range.y
    const end = Math.min(range.y + range.height, height)
    if (end - start < 2) return 'sorted'
    // Whole rows move, so a merged block in any of them would be torn apart; Excel refuses too.
    if (this.#layout(sheetId).merges.some(merge => merge.y < end && start < merge.y + merge.height)) return 'merged'
    const rows = Array.from({ length: end - start }, (_, index) => start + index)
    const values = new Map(rows.map(row => [row, this.#engine.getCellValue({ sheet: sheetId, col: byColumn, row })]))
    const sorted = [...rows].sort((left, right) => compareForSort(values.get(left) ?? null, values.get(right) ?? null, ascending))
    if (sorted.every((row, index) => row === rows[index])) return 'sorted'
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
    return 'sorted'
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

  selectionStats(sheetId: number, areas: readonly CellRange[]): SelectionStats {
    const { width, height } = this.#engine.getSheetDimensions(sheetId)
    let count = 0
    let numericCount = 0
    let sum = 0
    let min = Number.POSITIVE_INFINITY
    let max = Number.NEGATIVE_INFINITY
    for (const area of areas) {
      for (let row = area.y; row < Math.min(area.y + area.height, height); row += 1) {
        for (let col = area.x; col < Math.min(area.x + area.width, width); col += 1) {
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

  /**
   * The used area of a sheet, for exports: up to the last cell holding a value or formula. The engine's
   * own dimensions also reach every empty cell a formula reads (=SUM(A2:A1000) makes 1,000 rows).
   */
  usedSize(sheetId: number): { readonly width: number; readonly height: number } {
    let width = 0
    let height = 0
    this.#engine.getSheetSerialized(sheetId).forEach((row, rowIndex) => row.forEach((input, col) => {
      if (input === null || input === undefined || input === '') return
      width = Math.max(width, col + 1)
      height = rowIndex + 1
    }))
    return { width, height }
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

  #mapStyles(sheetId: number, areas: readonly CellRange[], map: (style: CellStyle) => CellStyle): void {
    const styles = this.#sheetStyles(this.#sheetName(sheetId))
    areas.forEach(area => this.#forEachCell(area, (col, row) => {
      const key = styleKey(col, row)
      const next = map(styles.get(key) ?? defaultCellStyle)
      if (isDefaultStyle(next)) styles.delete(key)
      else styles.set(key, next)
    }))
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

    if (axis === 'row') {
      const shiftedHeights = new Map<number, number>()
      layout.rowHeights.forEach((height, coordinate) => {
        if (removedCount > 0 && coordinate >= index && coordinate < index + removedCount) return
        shiftedHeights.set(coordinate >= index + removedCount ? coordinate + offset : coordinate, height)
      })
      layout.rowHeights.clear()
      shiftedHeights.forEach((height, coordinate) => layout.rowHeights.set(coordinate, height))
    }

    if (axis === 'column') {
      const shiftedWidths = new Map<number, number>()
      layout.columnWidths.forEach((width, coordinate) => {
        if (removedCount > 0 && coordinate >= index && coordinate < index + removedCount) return
        shiftedWidths.set(coordinate >= index + removedCount ? coordinate + offset : coordinate, width)
      })
      layout.columnWidths.clear()
      shiftedWidths.forEach((width, coordinate) => layout.columnWidths.set(coordinate, width))
    }

    layout.merges = layout.merges.flatMap(merge => {
      const shifted = shiftRange(merge, axis, index, offset, removedCount)
      return shifted === undefined || shifted.width * shifted.height < 2 ? [] : [shifted]
    })

    this.#namedRanges.forEach((namedRange, key) => {
      if (namedRange.sheetName !== sheetName) return
      const range = shiftRange(namedRange.range, axis, index, offset, removedCount)
      if (range === undefined) removedRangeKeys.push(key)
      else this.#namedRanges.set(key, { ...namedRange, range })
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
            rowHeights: [...layout.rowHeights.entries()],
            hiddenRows: [...layout.hiddenRows],
            hiddenColumns: [...layout.hiddenColumns],
            freezeRows: layout.freezeRows,
            freezeColumns: layout.freezeColumns,
            merges: layout.merges.map(cloneRange),
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
      // Excel formulas write TRUE and FALSE as plain words (VLOOKUP(…, FALSE)); the engine knows only TRUE() and FALSE().
      engine.addNamedExpression('TRUE', '=TRUE()')
      engine.addNamedExpression('FALSE', '=FALSE()')
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

  /** Folds a follow-up change into the undo step just committed, such as refitting rows after a style change. */
  #amend(mutation: () => void): void {
    mutation()
    this.#emit()
  }

  #emit(): void {
    this.#revision += 1
    this.#listeners.forEach(listener => listener())
  }
}
