import type { PDFPageProxy } from 'pdfjs-dist'
import { create } from 'zustand'
import type { Rect } from './engine/geometry'
import { isRedaction, newId, type PlacedMarkup } from './model'
import { buildPageText, evenAdvance, findMatches, rangeRects, textPieces, type Advance, type PageText } from './page-text'
import { commitPdf, readyPdf } from './pdf-store'
import { openPdfJs } from './pdfjs'

export interface FindMatch {
  readonly pageKey: string
  /** Where the match sits on the page as displayed, one rectangle per text item it spans. */
  readonly rects: readonly Rect[]
}

interface FindState {
  readonly open: boolean
  readonly query: string
  /** In page order, as far as the search has got. */
  readonly matches: readonly FindMatch[]
  /** Index into `matches` of the emphasised match. */
  readonly current: number
  readonly searching: boolean
  /** Bumped when Find is chosen while the bar is open, so the field takes focus again. */
  readonly focusRequest: number
}

export const useFindStore = create<FindState>(() => ({
  open: false,
  query: '',
  matches: [],
  current: 0,
  searching: false,
  focusRequest: 0,
}))

interface LoadedPage {
  readonly page: PDFPageProxy
  readonly text: PageText
}

/** Page text by source and page index; each source's pages keep their text while the source is open. */
const pageTexts = new WeakMap<Uint8Array, Map<number, Promise<LoadedPage>>>()

function loadPageText(source: Uint8Array, index: number): Promise<LoadedPage> {
  const pages = pageTexts.get(source) ?? new Map<number, Promise<LoadedPage>>()
  pageTexts.set(source, pages)
  const cached = pages.get(index)
  if (cached !== undefined) return cached
  const loading = (async () => {
    const page = await (await openPdfJs(source)).getPage(index + 1)
    return { page, text: buildPageText(textPieces(await page.getTextContent())) }
  })()
  // A failed read is tried again by the next search.
  void loading.catch(() => pages.delete(index))
  pages.set(index, loading)
  return loading
}

let measuring: OffscreenCanvasRenderingContext2D | null | undefined

/**
 * Places characters with the widths of the font pdf.js matched, scaled to the item's real width:
 * much closer than equal widths along a long line, which matters for redaction boxes.
 */
const measuredAdvance: Advance = (piece, offset) => {
  measuring ??= new OffscreenCanvas(1, 1).getContext('2d')
  if (measuring === null || offset === 0) return evenAdvance(piece, offset)
  measuring.font = `100px ${piece.fontFamily}`
  const whole = measuring.measureText(piece.str).width
  return whole > 0 ? measuring.measureText(piece.str.slice(0, offset)).width / whole : evenAdvance(piece, offset)
}

function reveal(): void {
  // After React has drawn the new current match.
  requestAnimationFrame(() => {
    window.document.querySelector('.pdf-find-hit.is-current')?.scrollIntoView({ block: 'center', inline: 'nearest' })
  })
}

/** The latest search; an older one stops when it sees it is no longer the latest. */
let generation = 0
/** How often a running search shows what it has found so far, in milliseconds. */
const publishInterval = 150

/**
 * Searches the pages in their current order, including pages inserted from other files. Matches
 * appear as they are found. With `jump`, the first match on or after the current page becomes
 * current and scrolls into view, as when typing a new query.
 */
export async function findInDocument(documentId: string, query: string, jump: boolean): Promise<void> {
  const run = (generation += 1)
  const document = readyPdf(documentId)
  const blank = query.trim() === ''
  useFindStore.setState(state => ({ query, matches: [], current: jump ? 0 : state.current, searching: !blank && document !== undefined }))
  if (blank || document === undefined) return

  const matches: FindMatch[] = []
  let chosen: number | undefined
  let published = performance.now()
  const publish = (searching: boolean): void => {
    const current = chosen ?? (jump ? 0 : Math.min(useFindStore.getState().current, Math.max(0, matches.length - 1)))
    useFindStore.setState({ matches: [...matches], current, searching })
    published = performance.now()
  }

  for (const [position, item] of document.present.pages.entries()) {
    // Stop once a newer search runs or the document has closed. Closing destroys the document's
    // pdf.js copy, and loading a page would open a new one, with its own worker, that nothing closes.
    // (The document's state changes identity on every edit, so only its absence means closed.)
    if (run !== generation || readyPdf(documentId) === undefined) return
    const source = document.sources[item.source]
    let found: FindMatch[] = []
    try {
      if (source === undefined) continue
      const { page, text } = await loadPageText(source, item.index)
      const viewport = page.getViewport({ scale: 1, rotation: (page.rotate + item.rotation) % 360 }).transform
      found = findMatches(text.text, query).map(range => ({ pageKey: item.key, rects: rangeRects(text, range, viewport, measuredAdvance) }))
    } catch {
      // A page whose text cannot be read has no matches.
    }
    if (run !== generation) return
    const revealNow = jump && chosen === undefined && position >= document.currentPage && found.length > 0
    if (revealNow) chosen = matches.length
    matches.push(...found)
    if (revealNow || performance.now() - published > publishInterval) publish(true)
    if (revealNow) reveal()
  }
  const wrapped = jump && chosen === undefined && matches.length > 0
  publish(false)
  if (wrapped) reveal()
}

export function openFind(): void {
  useFindStore.setState(state => ({ open: true, focusRequest: state.focusRequest + 1 }))
}

/** Hides the bar and its highlights; the query stays for next time. */
export function closeFind(): void {
  generation += 1
  useFindStore.setState({ open: false, matches: [], current: 0, searching: false })
}

export function stepFind(direction: 1 | -1): void {
  const { matches, current } = useFindStore.getState()
  if (matches.length === 0) return
  useFindStore.setState({ current: (current + direction + matches.length) % matches.length })
  reveal()
}

/** Room around each match so the box covers the glyphs whole; small enough not to reach the next line. */
const redactionPad = 1

function covers(outer: Rect, inner: Rect): boolean {
  const slack = 0.01
  return outer.x <= inner.x + slack && outer.y <= inner.y + slack
    && outer.x + outer.width >= inner.x + inner.width - slack && outer.y + outer.height >= inner.y + inner.height - slack
}

/**
 * Covers every match with a redaction box, as one undoable edit; saving applies them as it does
 * boxes drawn by hand. Matches a box already covers are left alone. Returns how many matches got boxes.
 */
export function redactMatches(documentId: string): number {
  const { matches } = useFindStore.getState()
  let marked = 0
  commitPdf(documentId, snapshot => {
    const pages = snapshot.pages.map(page => {
      const boxes: PlacedMarkup[] = []
      for (const match of matches) {
        if (match.pageKey !== page.key) continue
        const added = match.rects
          .map(rect => ({ x: rect.x - redactionPad, y: rect.y - redactionPad, width: rect.width + 2 * redactionPad, height: rect.height + 2 * redactionPad }))
          .filter(box => !page.markups.some(placed => isRedaction(placed.markup) && covers(placed.markup, box)))
        if (added.length > 0) marked += 1
        boxes.push(...added.map(box => ({ id: newId(), markup: { kind: 'redact' as const, ...box } })))
      }
      return boxes.length === 0 ? page : { ...page, markups: [...page.markups, ...boxes] }
    })
    return marked === 0 ? snapshot : { ...snapshot, pages }
  })
  return marked
}
