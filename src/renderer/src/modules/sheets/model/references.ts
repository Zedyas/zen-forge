import { columnIndex, columnName } from './address'

/** Excel's grid: references that move past it become #REF!. */
const lastColumn = 16_383
const lastRow = 1_048_575

/**
 * A1 references outside text and quoted sheet names: a cell or cell range, a column range (A:C) or a
 * row range (1:3). The lookarounds leave out function names (LOG10), names and number literals (1E5).
 */
const reference = /(?<![\w.$])(?:(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})|(\$?)(\d+):(\$?)(\d+))(?![\w(!.])/g

/** Text in quotes ("a""b") or a quoted sheet name ('Q1 ''26'), which never hold references to move. */
const quoted = /"(?:[^"]|"")*"?|'(?:[^']|'')*'?/g

function moveColumn(absolute: string, letters: string, columns: number): string | undefined {
  const index = columnIndex(letters) + (absolute === '$' ? 0 : columns)
  return index < 0 || index > lastColumn ? undefined : `${absolute}${columnName(index)}`
}

function moveRow(absolute: string, digits: string, rows: number): string | undefined {
  const index = Number(digits) - 1 + (absolute === '$' ? 0 : rows)
  return index < 0 || index > lastRow ? undefined : `${absolute}${index + 1}`
}

function moveReferences(text: string, columns: number, rows: number): string {
  return text.replace(reference, (match, ...groups: (string | undefined)[]) => {
    const [columnMark, column, rowMark, row, endColumnMark, endColumn, endRowMark, endRow,
      fromColumnMark, fromColumn, toColumnMark, toColumn, fromRowMark, fromRow, toRowMark, toRow] = groups
    const parts = column !== undefined
      ? [
          moveColumn(columnMark ?? '', column, columns),
          moveRow(rowMark ?? '', row ?? '', rows),
          ...(endColumn === undefined ? [] : [':', moveColumn(endColumnMark ?? '', endColumn, columns), moveRow(endRowMark ?? '', endRow ?? '', rows)]),
        ]
      : fromColumn !== undefined
        ? [moveColumn(fromColumnMark ?? '', fromColumn, columns), ':', moveColumn(toColumnMark ?? '', toColumn ?? '', columns)]
        : fromRow !== undefined
          ? [moveRow(fromRowMark ?? '', fromRow, rows), ':', moveRow(toRowMark ?? '', toRow ?? '', rows)]
          : [match]
    return parts.some(part => part === undefined) ? '#REF!' : parts.join('')
  })
}

/**
 * A formula copied `columns` across and `rows` down, as Copy and Paste or the fill handle write it:
 * relative references move by that distance, the `$` parts of a reference stay.
 */
export function shiftFormula(formula: string, columns: number, rows: number): string {
  if ((columns === 0 && rows === 0) || !formula.startsWith('=')) return formula
  let result = ''
  let last = 0
  for (const match of formula.matchAll(quoted)) {
    result += moveReferences(formula.slice(last, match.index), columns, rows) + match[0]
    last = match.index + match[0].length
  }
  return result + moveReferences(formula.slice(last), columns, rows)
}
