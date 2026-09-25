import type { Item } from '@glideapps/glide-data-grid'
import type { CellRange } from './workbook-data'

export function columnName(index: number): string {
  let value = index + 1
  let result = ''

  while (value > 0) {
    const remainder = (value - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    value = Math.floor((value - 1) / 26)
  }

  return result
}

export function cellName([column, row]: Item): string {
  return `${columnName(column)}${row + 1}`
}

/** `'Sales'!$A$1:$B$4`, the form named ranges use. The sheet name is always quoted, which every reader accepts. */
export function sheetRangeReference(sheetName: string, { x, y, width, height }: CellRange): string {
  const start = `$${columnName(x)}$${y + 1}`
  const end = `$${columnName(x + width - 1)}$${y + height}`
  return `'${sheetName.replaceAll("'", "''")}'!${start === end ? start : `${start}:${end}`}`
}

export function formulaReferences(formula: string): readonly string[] {
  return [...new Set(
    formula.match(/(?:'[^']+'|[A-Za-z_][\w.]*)?!?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/g) ?? [],
  )]
}
