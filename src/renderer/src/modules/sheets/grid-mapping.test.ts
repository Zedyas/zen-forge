import { describe, expect, it } from 'vitest'
import { coverBlocks, fillLines, mapRange, snapCell, visibleAreas } from './grid-mapping'

describe('visibleAreas', () => {
  // Rows 2 and 3 and column B are hidden.
  const rows = [0, 1, 4, 5, 6]
  const columns = [0, 2, 3]

  it('splits a view rectangle around hidden rows and columns', () => {
    expect(visibleAreas({ x: 0, y: 0, width: 3, height: 4 }, columns, rows)).toEqual([
      { x: 0, y: 0, width: 1, height: 2 },
      { x: 2, y: 0, width: 2, height: 2 },
      { x: 0, y: 4, width: 1, height: 2 },
      { x: 2, y: 4, width: 2, height: 2 },
    ])
    // The bounding range still spans the hidden rows, for row and column actions.
    expect(mapRange({ x: 0, y: 0, width: 3, height: 4 }, columns, rows)).toEqual({ x: 0, y: 0, width: 4, height: 6 })
  })

  it('keeps a selection without hidden rows as one rectangle', () => {
    expect(visibleAreas({ x: 1, y: 2, width: 2, height: 3 }, columns, rows)).toEqual([{ x: 2, y: 4, width: 2, height: 3 }])
  })
})

describe('fillLines', () => {
  it('reads a drag to the left as one backward line per row, nearest cell first', () => {
    const indexes = [0, 1, 2, 3]
    expect(fillLines({ x: 2, y: 1, width: 1, height: 1 }, { x: 0, y: 1, width: 3, height: 1 }, indexes, indexes)).toEqual({
      backward: true,
      lines: [{ seeds: [[2, 1]], targets: [[1, 1], [0, 1]] }],
    })
  })
})

describe('merged block selection', () => {
  // B2:C3 and E2:E5 are merged.
  const blocks = [{ x: 1, y: 1, width: 2, height: 2 }, { x: 4, y: 1, width: 1, height: 4 }]

  it('grows a range over every block it touches', () => {
    expect(coverBlocks({ x: 2, y: 2, width: 3, height: 1 }, blocks)).toEqual({ x: 1, y: 1, width: 4, height: 4 })
  })

  it('moves a click to the block’s first cell, and steps out of a block with the arrow keys', () => {
    expect(snapCell([2, 2], [0, 0], false, blocks)).toEqual([1, 1])
    expect(snapCell([1, 2], [1, 1], true, blocks)).toEqual([1, 3])
    expect(snapCell([3, 1], [2, 1], true, blocks)).toEqual([3, 1])
    expect(snapCell([2, 1], [1, 1], true, blocks)).toEqual([3, 1])
  })
})
