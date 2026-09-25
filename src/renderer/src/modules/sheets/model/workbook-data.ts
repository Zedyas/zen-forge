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
  /** `#rrggbb` */
  readonly textColor?: string
  /** `#rrggbb` */
  readonly fillColor?: string
  readonly align?: HorizontalAlign
  /**
   * The number format pattern from the opened file (`[$€-2] #,##0.00`, `dd/mm/yyyy`…). It is written
   * back unchanged, so a round trip keeps the file's exact format; editing the number format drops it.
   */
  readonly formatCode?: string
}

export const defaultCellStyle: CellStyle = { numberFormat: 'general', decimalPlaces: 2 }

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
  readonly hiddenRows: readonly number[]
  readonly hiddenColumns: readonly number[]
  readonly freezeRows: number
  readonly freezeColumns: number
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
