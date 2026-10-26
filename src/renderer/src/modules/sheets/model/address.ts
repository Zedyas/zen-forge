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

const sheetNameLimit = 31

/**
 * Why Excel would refuse `name` for a sheet, as a message for the person renaming it, or undefined when
 * it is allowed. `others` are the workbook's other sheet names; names match regardless of case.
 */
export function sheetNameProblem(name: string, others: readonly string[]): string | undefined {
  if (name === '') return 'A sheet name can’t be empty.'
  if (name.length > sheetNameLimit) return `A sheet name can be at most ${sheetNameLimit} characters.`
  if (/[/\\?*[\]:]/.test(name)) return 'A sheet name can’t contain / \\ ? * [ ] or :'
  if (name.startsWith("'") || name.endsWith("'")) return 'A sheet name can’t start or end with an apostrophe.'
  if (name.toLowerCase() === 'history') return 'Excel keeps the name History for itself.'
  const lower = name.toLocaleLowerCase()
  if (others.some(other => other.toLocaleLowerCase() === lower)) return `There is already a sheet named ${name}.`
  return undefined
}

/** A name Excel accepts, made from any text such as a file name: forbidden characters become `_`. */
export function sheetNameFrom(text: string): string {
  const name = text.replace(/[/\\?*[\]:]/g, '_').slice(0, sheetNameLimit).replace(/^'+|'+$/g, '').trim()
  return name === '' || name.toLowerCase() === 'history' ? 'Sheet1' : name
}
