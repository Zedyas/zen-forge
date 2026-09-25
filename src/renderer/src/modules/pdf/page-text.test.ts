import { describe, expect, it } from 'vitest'
import type { Rect } from './engine/geometry'
import { buildPageText, findMatches, rangeRects, type TextPiece } from './page-text'

/** 10pt text with every character 5pt wide, at a baseline origin in user space. */
function piece(str: string, x: number, y: number, hasEOL = false): TextPiece {
  return { str, transform: [10, 0, 0, 10, x, y], width: str.length * 5, hasEOL, ascent: 0.8, descent: -0.2, fontFamily: 'sans-serif' }
}

function rounded(rects: readonly Rect[]): Rect[] {
  return rects.map(rect => ({ x: +rect.x.toFixed(3), y: +rect.y.toFixed(3), width: +rect.width.toFixed(3), height: +rect.height.toFixed(3) }))
}

describe('page text', () => {
  // Two lines of a 612 × 792 page. pdf.js marks the first line's end; a gap separates "due" and "now".
  const page = buildPageText([piece('Monthly ', 100, 700), piece('rent', 140, 700, true), piece('is due', 100, 686), piece('now', 134, 686)])

  it('joins items with a line break at the end of a line and a space at a gap', () => {
    expect(page.text).toBe('Monthly rent\nis due now')
  })

  it('maps a case-insensitive match spanning two items to one rectangle per item, in displayed coordinates', () => {
    const matches = findMatches(page.text, 'RENT IS')
    expect(matches).toEqual([{ start: 8, end: 15 }])
    const [match] = matches
    if (match === undefined) return

    // pdf.js's viewport for the unrotated page flips y: user (x, y) is displayed at (x, 792 - y).
    expect(rounded(rangeRects(page, match, [1, 0, 0, -1, 0, 792]))).toEqual([
      { x: 140, y: 84, width: 20, height: 10 },
      { x: 100, y: 98, width: 10, height: 10 },
    ])
    // Turned a quarter clockwise, user (x, y) is displayed at (y, x): the boxes turn with the text.
    expect(rounded(rangeRects(page, match, [0, 1, 1, 0, 0, 0]))).toEqual([
      { x: 698, y: 140, width: 10, height: 20 },
      { x: 684, y: 100, width: 10, height: 10 },
    ])
  })
})
