import type { Item, Rectangle } from '@glideapps/glide-data-grid'
import { useSyncExternalStore } from 'react'
import { toast } from 'sonner'
import { z } from 'zod'
import { delimitedText, readTabbedText } from './io'
import { parseCellText } from './model/cell-parse'
import { formatCellValue } from './model/format'
import type { CellRange, CellStyle, CopiedCells, PastedCell, Workbook } from './model/Workbook'

/*
 * Edit > Copy, Cut and Paste reach the page as clipboard events. Ledger puts two things on the
 * clipboard: its own copy of the cells in a private type (raw inputs, so formulas stay formulas, and
 * styles), which pastes into any Ledger tab or window, and the shown values as tab-separated text for
 * other apps. A paste from another app reads that text.
 */

const clipboardType = 'application/x-zen-suzu-ledger+json'

/* Any app can write the private type, so a paste accepts only JSON with exactly the model's shape. */

const color = z.string().regex(/^#[0-9a-f]{6}$/i)
const styleFields = {
  numberFormat: z.enum(['general', 'number', 'currency', 'percent', 'date']),
  decimalPlaces: z.number().int().min(0).max(30),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  strike: z.boolean().optional(),
  wrap: z.boolean().optional(),
  fontFamily: z.string().max(100).optional(),
  fontSize: z.number().min(1).max(409).optional(),
  textColor: color.optional(),
  fillColor: color.optional(),
  align: z.enum(['left', 'center', 'right']).optional(),
  formatCode: z.string().max(255).optional(),
} satisfies Record<keyof CellStyle, z.ZodType>
const index = z.number().int().min(0).max(1_048_575)
const cell = z.object({
  input: z.union([z.string().max(32_767), z.number().finite(), z.boolean(), z.null()]),
  style: z.object(styleFields).optional(),
})
const payload = z.object({
  columns: z.array(index).min(1).max(16_384),
  rows: z.array(index).min(1).max(1_048_576),
  cells: z.array(z.array(cell)),
  cut: z.string().optional(),
}).refine(value => value.cells.length === value.rows.length && value.cells.every(row => row.length === value.columns.length))

type ClipboardPayload = CopiedCells & { readonly cut?: string }

function readPayload(json: string): ClipboardPayload | undefined {
  try {
    const parsed = payload.safeParse(JSON.parse(json))
    return parsed.success ? parsed.data : undefined
  } catch {
    return undefined
  }
}

/** Cut cells wait for the paste, as in Excel: the paste moves them, so formulas that refer to them follow. */
interface PendingCut {
  readonly id: string
  readonly workbook: Workbook
  readonly sheetId: number
  readonly range: CellRange
  /** The workbook's revision at the cut; any later edit turns the cut into a plain copy. */
  readonly revision: number
}

let pendingCut: PendingCut | undefined
const listeners = new Set<() => void>()

function setPendingCut(next: PendingCut | undefined): void {
  pendingCut = next
  listeners.forEach(listener => listener())
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

function livePendingCut(): PendingCut | undefined {
  return pendingCut !== undefined && pendingCut.revision === pendingCut.workbook.getSnapshot() ? pendingCut : undefined
}

/** The cut range waiting on this sheet, for the grid's dashed outline. */
export function usePendingCut(workbook: Workbook, sheetId: number): CellRange | undefined {
  const cut = useSyncExternalStore(subscribe, () => pendingCut)
  useSyncExternalStore(workbook.subscribe, workbook.getSnapshot)
  return cut !== undefined && cut === livePendingCut() && cut.workbook === workbook && cut.sheetId === sheetId ? cut.range : undefined
}

export function cancelCut(): void {
  if (pendingCut !== undefined) setPendingCut(undefined)
}

export interface CopySource {
  readonly workbook: Workbook
  readonly sheetId: number
  /** The selection's visible columns and rows, in view order. */
  readonly columns: readonly number[]
  readonly rows: readonly number[]
  /** The selection from first to last cell; a cut moves all of it. */
  readonly range: CellRange
}

export function copyCells(data: DataTransfer, source: CopySource, cut: boolean): void {
  const { workbook, sheetId, columns, rows } = source
  const id = cut ? crypto.randomUUID() : undefined
  const copied: ClipboardPayload = { ...workbook.copyCells(sheetId, columns, rows), ...(id === undefined ? {} : { cut: id }) }
  data.setData(clipboardType, JSON.stringify(copied))
  data.setData('text/plain', delimitedText(rows.map(row => columns.map(col =>
    formatCellValue(workbook.getCellValue(sheetId, [col, row]), workbook.getStyle(sheetId, [col, row])))), 'tsv'))
  setPendingCut(id === undefined ? undefined : { id, workbook, sheetId, range: source.range, revision: workbook.getSnapshot() })
}

/**
 * Where each cell of a `height` by `width` block lands when pasted on `selection` (view coordinates):
 * view cells map to workbook cells, so hidden rows are skipped. A selection that is a whole multiple
 * of the block is filled with copies of it, as in Excel. Cells past the grid's edge are left out.
 */
export function pasteTargets(
  height: number,
  width: number,
  selection: Rectangle,
  columns: readonly number[],
  rows: readonly number[],
): readonly { readonly target: Item; readonly row: number; readonly column: number }[] {
  const tiles = selection.width % width === 0 && selection.height % height === 0
  const spanRows = tiles ? selection.height : height
  const spanColumns = tiles ? selection.width : width
  const result: { target: Item; row: number; column: number }[] = []
  for (let viewRow = 0; viewRow < spanRows; viewRow += 1) {
    const row = rows[selection.y + viewRow]
    if (row === undefined) break
    for (let viewColumn = 0; viewColumn < spanColumns; viewColumn += 1) {
      const col = columns[selection.x + viewColumn]
      if (col === undefined) break
      result.push({ target: [col, row], row: viewRow % height, column: viewColumn % width })
    }
  }
  return result
}

export interface PasteTarget {
  readonly workbook: Workbook
  readonly sheetId: number
  /** The selection in view coordinates, and each view column's and row's workbook index. */
  readonly selection: Rectangle
  readonly columns: readonly number[]
  readonly rows: readonly number[]
}

function copiedCells(copied: CopiedCells, target: PasteTarget): PastedCell[] {
  return pasteTargets(copied.rows.length, copied.columns.length, target.selection, target.columns, target.rows)
    .flatMap(({ target: at, row, column }) => {
      const source = copied.cells[row]?.[column]
      const sourceCol = copied.columns[column]
      const sourceRow = copied.rows[row]
      if (source === undefined || sourceCol === undefined || sourceRow === undefined) return []
      return [{ target: at, input: source.input, style: source.style, source: [sourceCol, sourceRow] as const }]
    })
}

function textCells(text: string, target: PasteTarget): PastedCell[] {
  const fields = readTabbedText(text)
  const width = Math.max(0, ...fields.map(row => row.length))
  if (fields.length === 0 || width === 0) return []
  return pasteTargets(fields.length, width, target.selection, target.columns, target.rows).flatMap(({ target: at, row, column }) => {
    const field = fields[row]?.[column]
    return field === undefined ? [] : [{ target: at, ...parseCellText(field) }]
  })
}

/** Pastes Ledger's own cells when the clipboard has them, otherwise tab-separated text. False when it holds neither. */
export function pasteCells(data: DataTransfer, target: PasteTarget): boolean {
  const json = data.getData(clipboardType)
  if (json !== '') {
    const copied = readPayload(json)
    if (copied === undefined) {
      toast('Nothing pasted', { description: 'The copied cells are not ones Ledger can use.' })
      return true
    }
    const cut = livePendingCut()
    if (cut !== undefined && copied.cut === cut.id) {
      setPendingCut(undefined)
      const destination = [target.columns[target.selection.x] ?? 0, target.rows[target.selection.y] ?? 0] as const
      if (cut.workbook === target.workbook) {
        if (!target.workbook.moveCells(cut.sheetId, cut.range, target.sheetId, destination)) {
          toast.error('Those cells can’t move there', { description: 'Part of an array formula would be split.' })
        }
        return true
      }
      // Into another workbook, the cells are copied and then cleared where they were.
      target.workbook.pasteCells(target.sheetId, copiedCells(copied, target))
      cut.workbook.clearCells(cut.sheetId, [cut.range])
      return true
    }
    target.workbook.pasteCells(target.sheetId, copiedCells(copied, target))
    return true
  }
  const text = data.getData('text/plain')
  if (text === '') return false
  target.workbook.pasteCells(target.sheetId, textCells(text, target))
  return true
}
