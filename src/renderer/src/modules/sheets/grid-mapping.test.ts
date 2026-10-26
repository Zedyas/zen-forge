import { describe, expect, it } from 'vitest'
import { mapRange, visibleAreas } from './grid-mapping'

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
