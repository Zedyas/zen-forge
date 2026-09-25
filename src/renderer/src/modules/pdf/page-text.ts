import type { PDFPageProxy } from 'pdfjs-dist'
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
  /** pdf.js's writing direction: `ltr`, `rtl`, or `ttb` for vertical text. */
  readonly dir: string
  /** Text space to PDF user space, with the font size folded in: [a, b, c, d, e, f], origin on the baseline. */
  readonly transform: readonly number[]
  /** Advance along the baseline, in user space. */
  readonly width: number
  /** For vertical text, the advance down the column, in user space. */
  readonly height: number
  /** pdf.js marks the last item of a line. */
  readonly hasEOL: boolean
  /** Font ascent and descent as fractions of the font size; descent is negative. */
  readonly ascent: number
  readonly descent: number
  /** The CSS font family pdf.js matches the font to, for measuring characters. */
  readonly fontFamily: string
}

/** What pdf.js's `getTextContent()` returns. */
type TextContent = Awaited<ReturnType<PDFPageProxy['getTextContent']>>

/** The text items of a page, with their fonts' ascent and descent (pdf.js leaves them 0 when unknown). */
export function textPieces(content: TextContent): TextPiece[] {
  return content.items.flatMap(item => {
    if (!('str' in item)) return []
    const style = content.styles[item.fontName]
    const ascent = style !== undefined && style.ascent > 0 ? style.ascent : 0.8
    const descent = style !== undefined && style.descent < 0 ? style.descent : ascent - 1
    const { str, dir, transform, width, height, hasEOL } = item
    return [{ str, dir, transform, width, height, hasEOL, ascent, descent, fontFamily: style?.fontFamily ?? 'sans-serif' }]
  })
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

/** Characters pdf.js reorders into right-to-left (logical) order when it builds an item's text. */
const rightToLeft = /[\u0590-\u08ff\ufb1d-\ufdff\ufe70-\ufefc]/

/**
 * The part of a piece that characters `from` to `to` cover, in the piece's text space (in ems: x
 * along the baseline, y up from it), as [left, right, bottom, top]. Only left-to-right text is cut
 * by character. pdf.js gives right-to-left text in reading order but places it from the left, so
 * measuring from the left would cover the wrong letters: a range there covers the whole piece.
 * So does a range in vertical text, whose glyphs hang below the text position, centred on it.
 */
function textSpaceBox(piece: TextPiece, from: number, to: number, advance: Advance): readonly [number, number, number, number] {
  const [a = 1, b = 0, c = 0, d = 1] = piece.transform
  if (piece.dir === 'ttb') return [-0.5, 0.5, -piece.height / (Math.hypot(c, d) || 1), 0]
  const length = piece.width / (Math.hypot(a, b) || 1)
  if (piece.dir !== 'ltr' || rightToLeft.test(piece.str)) return [0, length, piece.descent, piece.ascent]
  return [length * advance(piece, from), length * advance(piece, to), piece.descent, piece.ascent]
}

/**
 * The rectangles a range of the page text covers, one per piece it touches, in the coordinates
 * `viewport` maps user space to (pdf.js's viewport transform: displayed page points, top-left
 * origin, rotation applied). Each spans the font's ascent to its descent, and a piece that does not
 * run left to right is covered whole.
 */
export function rangeRects(page: PageText, range: TextRange, viewport: readonly number[], advance: Advance = evenAdvance): Rect[] {
  const [m0 = 1, m1 = 0, m2 = 0, m3 = 1, m4 = 0, m5 = 0] = viewport
  return page.spans.flatMap(({ piece, start }) => {
    const from = Math.max(range.start, start) - start
    const to = Math.min(range.end, start + piece.str.length) - start
    if (from >= to) return []
    const [a = 1, b = 0, c = 0, d = 1, e = 0, f = 0] = piece.transform
    const [x0, x1, y0, y1] = textSpaceBox(piece, from, to, advance)
    const corners = [x0, x1].flatMap(x => [y0, y1].map(y => {
      // Text space to user space, then user space to the page as displayed.
      const userX = a * x + c * y + e
      const userY = b * x + d * y + f
      return [m0 * userX + m2 * userY + m4, m1 * userX + m3 * userY + m5] as const
    }))
    const xs = corners.map(([x]) => x)
    const ys = corners.map(([, y]) => y)
    const left = Math.min(...xs)
    const top = Math.min(...ys)
    return [{ x: left, y: top, width: Math.max(...xs) - left, height: Math.max(...ys) - top }]
  })
}
