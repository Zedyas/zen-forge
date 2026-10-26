import type { Item } from '@glideapps/glide-data-grid'
import { toast } from 'sonner'
import { askConfirm } from '../../ui/ConfirmDialog'
import { measureCellHeight } from './cell-text'
import type { CellRange, CellStyle, Workbook } from './model/Workbook'

/** The active grid's state that toolbar buttons and menu commands act on. */
export interface SheetController {
  readonly documentId: string
  readonly workbook: Workbook
  readonly sheetId: number
  readonly sheetName: string
  /** The selection from its first to last cell, for row, column and sort actions. */
  readonly range: CellRange
  /** The selection's visible cells as rectangles, for cell-by-cell actions such as formatting. */
  readonly areas: readonly CellRange[]
  readonly cell: Item
  openFind(): void
}

let active: SheetController | undefined

/** Registers the mounted grid as the target of Format and Edit menu commands; returns the unregisterer. */
export function setActiveSheet(controller: SheetController): () => void {
  active = controller
  return () => {
    if (active === controller) active = undefined
  }
}

export function activeSheet(): SheetController | undefined {
  return active
}

type ToggleKey = 'bold' | 'italic' | 'underline' | 'strike' | 'wrap'

/** Toggles by the active cell, like Excel and Numbers: bold active cell → the whole range loses bold. */
export function toggleStyle(controller: SheetController, key: ToggleKey): void {
  const on = controller.workbook.getStyle(controller.sheetId, controller.cell)[key] === true
  setStyle(controller, { [key]: on ? undefined : true })
}

/** Rows refit to their text when wrapping or the font changes. */
export function setStyle(controller: SheetController, patch: Partial<CellStyle>): void {
  controller.workbook.updateStyle(controller.sheetId, controller.areas, patch, measureCellHeight)
}

/** The selection's visible rows, for row height. */
export function selectedRows(controller: SheetController): readonly number[] {
  return [...new Set(controller.areas.flatMap(area => Array.from({ length: area.height }, (_, offset) => area.y + offset)))]
}

/** Font families on every Mac; a file's other fonts are kept and listed when a cell uses one. */
export const fontFamilies = ['Helvetica Neue', 'Arial', 'Times New Roman', 'Georgia', 'Courier New', 'Menlo'] as const
export const fontSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72] as const

/**
 * Sorts the selected rows by the active cell's column. With a single cell selected it sorts the
 * sheet's data below a detected header row, which is what "sort by this column" usually means.
 */
export function sortByActiveColumn(controller: SheetController, ascending: boolean): void {
  const { workbook, sheetId, range, cell } = controller
  const target = range.height > 1 ? range : workbook.sortRegion(sheetId)
  if (target === undefined) {
    toast.info('Nothing to sort', { description: 'Select at least two rows, or add data below the header.' })
    return
  }
  if (workbook.sortRows(sheetId, target, cell[0], ascending) === 'merged') {
    toast.info('Merged cells can’t be sorted', { description: 'Unmerge the cells in these rows, then sort.' })
  }
}

/** Merges the selection, or each of its rows with `across`; asks first when values would be discarded. */
export async function mergeSelection(controller: SheetController, across: boolean): Promise<void> {
  const { workbook, sheetId, range } = controller
  if (range.width * range.height < 2) {
    toast.info('Select more than one cell to merge')
    return
  }
  const discarded = workbook.mergeDiscards(sheetId, range, across)
  if (discarded > 0) {
    const { value } = await askConfirm({
      title: 'Merge cells?',
      message: `Merging keeps only the upper-left value${across ? ' of each row' : ''} and discards the other ${discarded === 1 ? 'value' : `${discarded} values`}.`,
      actions: [{ label: 'Merge', value: 'merge', primary: true }],
    })
    if (value !== 'merge') return
  }
  workbook.mergeCells(sheetId, range, across)
}

export function autoSum(controller: SheetController): void {
  if (!controller.workbook.autoSum(controller.sheetId, controller.cell)) {
    toast.info('Nothing to sum', { description: 'AutoSum adds the numbers directly above or to the left of the selected cell.' })
  }
}

export const textColors: ReadonlyArray<{ readonly label: string; readonly value: string | undefined }> = [
  { label: 'Automatic', value: undefined },
  { label: 'Black', value: '#1d1f23' },
  { label: 'Grey', value: '#6b7280' },
  { label: 'Red', value: '#c9352b' },
  { label: 'Orange', value: '#c2620a' },
  { label: 'Green', value: '#1f8a4c' },
  { label: 'Blue', value: '#2f5bd3' },
  { label: 'Purple', value: '#7a3fc8' },
]

export const fillColors: ReadonlyArray<{ readonly label: string; readonly value: string | undefined }> = [
  { label: 'No fill', value: undefined },
  { label: 'Yellow', value: '#fdf1c7' },
  { label: 'Orange', value: '#fde2cc' },
  { label: 'Red', value: '#f8dde0' },
  { label: 'Purple', value: '#ebe4f7' },
  { label: 'Blue', value: '#dbe8fb' },
  { label: 'Green', value: '#dcf2e6' },
  { label: 'Grey', value: '#e9ebee' },
]
