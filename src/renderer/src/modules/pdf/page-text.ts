import type { Rect } from './engine/geometry'

/*
 * Page text for search. pdf.js splits a page's text into items, one per run of text drawn
 * together; a phrase can span several. The items join into one string per page, with a space or
 * a line break wherever the layout breaks between them, so a search runs over plain text and each
 * match maps back to the items (and the part of each item) it covers.
 */

/** What search needs from one pdf.js text item. */
export interface TextPiece {
  readonly str: string
  /** Text space to PDF user space, with the font size folded in: [a, b, c, d, e, f], origin on the baseline. */
  readonly transform: readonly number[]
  /** Advance along the baseline, in user space. */
  readonly width: number
  /** pdf.js marks the last item of a line. */
  readonly hasEOL: boolean
  /** Font ascent and descent as fractions of the font size; descent is negative. */
  readonly ascent: number
  readonly descent: number
  /** The CSS font family pdf.js matches the font to, for measuring characters. */
  readonly fontFamily: string
}

/** Where the character at `offset` starts along a piece, as a fraction of the piece's width. */
export type Advance = (piece: TextPiece, offset: number) => number

/** Every character equally wide: good enough to highlight, and what is left when nothing better is known. */
export const evenAdvance: Advance = (piece, offset) => offset / piece.str.length

export interface PageText {
  readonly text: string
  /** Each piece and the offset in `text` where its characters start. */
  readonly spans: readonly { readonly piece: TextPiece; readonly start: number }[]
}

export interface TextRange {
  readonly start: number
  readonly end: number
}

/** A line break when the next piece starts on another line, a space when it starts after a gap. */
function separator(previous: TextPiece, next: TextPiece): string {
  if (previous.hasEOL) return '\n'
  if (/\s$/.test(previous.str) || /^\s/.test(next.str)) return ''
  const [a = 1, b = 0, c = 0, d = 1, e = 0, f = 0] = previous.transform
  const scale = Math.hypot(a, b) || 1
  const size = Math.hypot(c, d) || scale
  // Offset from where the previous piece ends to where the next begins, along and across its baseline.
  const dx = (next.transform[4] ?? 0) - (e + (a / scale) * previous.width)
  const dy = (next.transform[5] ?? 0) - (f + (b / scale) * previous.width)
  const along = (dx * a + dy * b) / scale
  const across = (dy * a - dx * b) / scale
  if (Math.abs(across) > size / 2) return '\n'
  return along > size / 10 ? ' ' : ''
}

export function buildPageText(pieces: readonly TextPiece[]): PageText {
  let text = ''
  const spans: { piece: TextPiece; start: number }[] = []
  pieces.forEach((piece, index) => {
    const previous = pieces[index - 1]
    if (previous !== undefined) text += separator(previous, piece)
    spans.push({ piece, start: text.length })
    text += piece.str
  })
  return { text, spans }
}

/** Case-insensitive matches of `query`; any run of whitespace in it matches any run in the text, line breaks included. */
export function findMatches(text: string, query: string): TextRange[] {
  const words = query.trim().split(/\s+/).filter(word => word !== '')
  if (words.length === 0) return []
  const pattern = new RegExp(words.map(word => word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('\\s+'), 'giu')
  return Array.from(text.matchAll(pattern), match => ({ start: match.index, end: match.index + match[0].length }))
}

/**
 * The rectangles a range of the page text covers, one per piece it touches, in the coordinates
 * `viewport` maps user space to (pdf.js's viewport transform: displayed page points, top-left
 * origin, rotation applied). Each spans the font's ascent to its descent.
 */
export function rangeRects(page: PageText, range: TextRange, viewport: readonly number[], advance: Advance = evenAdvance): Rect[] {
  const [m0 = 1, m1 = 0, m2 = 0, m3 = 1, m4 = 0, m5 = 0] = viewport
  return page.spans.flatMap(({ piece, start }) => {
    const from = Math.max(range.start, start) - start
    const to = Math.min(range.end, start + piece.str.length) - start
    if (from >= to) return []
    const [a = 1, b = 0, c = 0, d = 1, e = 0, f = 0] = piece.transform
    const scale = Math.hypot(a, b) || 1
    const corners = [from, to].flatMap(offset => {
      const along = piece.width * advance(piece, offset)
      return [piece.descent, piece.ascent].map(up => {
        const x = e + (a / scale) * along + c * up
        const y = f + (b / scale) * along + d * up
        return [m0 * x + m2 * y + m4, m1 * x + m3 * y + m5] as const
      })
    })
    const xs = corners.map(([x]) => x)
    const ys = corners.map(([, y]) => y)
    const left = Math.min(...xs)
    const top = Math.min(...ys)
    return [{ x: left, y: top, width: Math.max(...xs) - left, height: Math.max(...ys) - top }]
  })
}
