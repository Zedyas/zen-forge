import type { Item, Rectangle } from '@glideapps/glide-data-grid'
import type { CellRange, FillLine } from './model/Workbook'

/*
 * The grid shows only visible rows and columns, so its (view) coordinates skip hidden ones. `columns`
 * and `rows` list the workbook index of each view index.
 */

export function mapCell([col, row]: Item, columns: readonly number[], rows: readonly number[]): Item {
  return [columns[col] ?? 0, rows[row] ?? 0]
}

/** The workbook rectangle from a view rectangle's first to last cell, including any hidden rows and columns between. */
export function mapRange(range: Rectangle, columns: readonly number[], rows: readonly number[]): CellRange {
  const first = mapCell([range.x, range.y], columns, rows)
  const last = mapCell([range.x + Math.max(1, range.width) - 1, range.y + Math.max(1, range.height) - 1], columns, rows)
  return {
    x: Math.min(first[0], last[0]),
    y: Math.min(first[1], last[1]),
    width: Math.abs(last[0] - first[0]) + 1,
    height: Math.abs(last[1] - first[1]) + 1,
  }
}

/** Runs of consecutive workbook indexes, as [first, length], for view indexes `start` to `start + count - 1`. */
function runs(indexes: readonly number[], start: number, count: number): readonly (readonly [number, number])[] {
  const result: [number, number][] = []
  for (let view = start; view < start + Math.max(1, count); view += 1) {
    const index = indexes[view]
    if (index === undefined) continue
    const last = result.at(-1)
    if (last !== undefined && last[0] + last[1] === index) last[1] += 1
    else result.push([index, 1])
  }
  return result
}

/**
 * The workbook rectangles a view rectangle covers, split around hidden rows and columns, so clearing,
 * formatting and statistics skip what is hidden, as Excel does with filtered rows.
 */
export function visibleAreas(range: Rectangle, columns: readonly number[], rows: readonly number[]): readonly CellRange[] {
  const columnRuns = runs(columns, range.x, range.width)
  return runs(rows, range.y, range.height).flatMap(([y, height]) =>
    columnRuns.map(([x, width]) => ({ x, y, width, height })))
}

/** First and last view index whose workbook index lies in `start` to `start + count - 1`. */
function viewSpan(indexes: readonly number[], start: number, count: number): readonly [number, number] | undefined {
  const first = indexes.findIndex(index => index >= start && index < start + count)
  if (first < 0) return undefined
  let last = first
  while ((indexes[last + 1] ?? Number.POSITIVE_INFINITY) < start + count) last += 1
  return [first, last]
}

/** The view rectangle showing a workbook range, or undefined when all of it is hidden. */
export function viewRange(range: CellRange, columns: readonly number[], rows: readonly number[]): Rectangle | undefined {
  const across = viewSpan(columns, range.x, range.width)
  const down = viewSpan(rows, range.y, range.height)
  if (across === undefined || down === undefined) return undefined
  return { x: across[0], y: down[0], width: across[1] - across[0] + 1, height: down[1] - down[0] + 1 }
}

/**
 * A fill-handle drag (view rectangles: the selection it started from, and the selection it ended
 * with) as lines of workbook cells, one per column for a vertical fill and one per row for a horizontal one.
 */
export function fillLines(
  source: Rectangle,
  destination: Rectangle,
  columns: readonly number[],
  rows: readonly number[],
): { readonly lines: readonly FillLine[]; readonly backward: boolean } {
  const vertical = destination.height !== source.height
  const backward = vertical ? destination.y < source.y : destination.x < source.x
  const span = (start: number, end: number): number[] => Array.from({ length: Math.max(0, end - start) }, (_, offset) => start + offset)
  const across = vertical ? span(source.x, source.x + source.width) : span(source.y, source.y + source.height)
  const seeds = vertical ? span(source.y, source.y + source.height) : span(source.x, source.x + source.width)
  const [start, end] = vertical ? [destination.y, destination.y + destination.height] : [destination.x, destination.x + destination.width]
  const sourceEnd = seeds.at(-1) ?? 0
  const targets = backward ? span(start, seeds[0] ?? 0).reverse() : span(sourceEnd + 1, end)
  const cell = (line: number, along: number): Item => mapCell(vertical ? [line, along] : [along, line], columns, rows)
  return {
    backward,
    lines: across.map(line => ({ seeds: seeds.map(along => cell(line, along)), targets: targets.map(along => cell(line, along)) })),
  }
}
