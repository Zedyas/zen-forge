/** Plain, serializable workbook contents: the one shape shared by the live Workbook and file formats. */

export type CellInput = string | number | boolean | null
export type NumberFormat = 'general' | 'number' | 'currency' | 'percent' | 'date'
export type HorizontalAlign = 'left' | 'center' | 'right'

export interface CellRange {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

/** Absent optional fields mean the default: regular weight, no colour, alignment by value type. */
export interface CellStyle {
  readonly numberFormat: NumberFormat
  readonly decimalPlaces: number
  readonly bold?: boolean
  readonly italic?: boolean
  readonly underline?: boolean
  readonly strike?: boolean
  /** Long text breaks onto more lines, and the row grows to show them. */
  readonly wrap?: boolean
  /** A font family name; absent means the default font. */
  readonly fontFamily?: string
  /** In points, as Excel counts them; absent means `defaultFontSize`. */
  readonly fontSize?: number
  /** `#rrggbb` */
  readonly textColor?: string
  /** `#rrggbb` */
  readonly fillColor?: string
  readonly align?: HorizontalAlign
  /**
   * An Excel number format code (`[$€-2] #,##0.00`, `dd/mm/yyyy`…), from the opened file or typed as a
   * custom format. Cells show with it when Ledger supports the code (format-code.ts), and it is written
   * back unchanged either way; choosing one of the five formats drops it.
   */
  readonly formatCode?: string
}

export const defaultCellStyle: CellStyle = { numberFormat: 'general', decimalPlaces: 2 }

/** Excel's default font size, in points; Ledger shows it at its standard cell text size. */
export const defaultFontSize = 11
/** In pixels. */
export const defaultRowHeight = 26

/** Every sheet shows at least this much grid, however little data it holds. */
export const defaultRowCount = 1_000
export const defaultColumnCount = 26

/** Whether a style adds nothing to the default, so the cell needs no stored style. */
export function isDefaultStyle(style: CellStyle): boolean {
  return (Object.keys(style) as (keyof CellStyle)[]).every(key =>
    style[key] === undefined || style[key] === defaultCellStyle[key])
}

export interface NamedRange {
  readonly name: string
  readonly sheetName: string
  readonly range: CellRange
}

export interface SheetLayoutData {
  readonly rowCount: number
  readonly columnCount: number
  /** [column, width in px] */
  readonly columnWidths: readonly (readonly [number, number])[]
  /** [row, height in px] for rows that are not `defaultRowHeight`. */
  readonly rowHeights: readonly (readonly [number, number])[]
  readonly hiddenRows: readonly number[]
  readonly hiddenColumns: readonly number[]
  readonly freezeRows: number
  readonly freezeColumns: number
  /** Merged blocks; each shows its top-left cell across the whole block. */
  readonly merges: readonly CellRange[]
}

export interface SheetData {
  readonly name: string
  /** Row-major raw inputs. A formula is a string starting with `=`. Ragged rows are allowed. */
  readonly cells: readonly (readonly CellInput[])[]
  /**
   * Row-major computed values, same shape as `cells`. Exporters write these as cached formula
   * results so a file opens with numbers visible before recalculation. Importers may omit it.
   */
  readonly values?: readonly (readonly (string | number | boolean | null)[])[]
  /** [column, row, style] for every cell whose style differs from `defaultCellStyle`. */
  readonly styles: readonly (readonly [number, number, CellStyle])[]
  readonly layout: SheetLayoutData
}

export interface WorkbookData {
  readonly sheets: readonly SheetData[]
  readonly namedRanges: readonly NamedRange[]
}

/** A copied cell: its raw input (a formula stays a formula) and its style, if any. */
export interface CopiedCell {
  readonly input: CellInput
  readonly style?: CellStyle
}

/** Copied cells, as they travel on the clipboard between Ledger tabs and windows. */
export interface CopiedCells {
  /** The workbook column and row of each copied column and row; hidden ones are left out, so they can skip numbers. */
  readonly columns: readonly number[]
  readonly rows: readonly number[]
  /** Row-major, `rows.length` by `columns.length`. */
  readonly cells: readonly (readonly CopiedCell[])[]
}

/** One cell a paste writes. */
export interface PastedCell {
  readonly target: readonly [number, number]
  readonly input: CellInput
  /** Replaces the target's style; absent means the default style. */
  readonly style?: CellStyle
  /** Where a copied formula came from: its relative references move by the distance to `target`. */
  readonly source?: readonly [number, number]
}

