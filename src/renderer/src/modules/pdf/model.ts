import type { PageRef, PageRotation, PdfEdit, PdfPoint } from './engine'

/**
 * Everything the user has changed about one PDF, as plain data. The original bytes are never
 * edited: a save replays this snapshot onto the sources, so undo is just keeping old snapshots.
 */

export interface RedactBox {
  readonly kind: 'redact'
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

/** An engine edit, or a redaction box (applied at save time by removing what lies under it). */
export type Markup = PdfEdit | RedactBox

export interface PlacedMarkup {
  readonly id: string
  readonly markup: Markup
}

export interface PageItem {
  /** Stable across reorders; used for thumbnails, drag and drop, and selection. */
  readonly key: string
  readonly source: number
  readonly index: number
  /** Clockwise rotation added by the user on top of the page's own /Rotate. */
  readonly rotation: PageRotation
  /** Coordinates in points, top-left origin, on the page as displayed after `rotation`. */
  readonly markups: readonly PlacedMarkup[]
}

export interface Snapshot {
  readonly pages: readonly PageItem[]
  readonly formValues: Readonly<Record<string, string | boolean>>
  readonly flattenForm: boolean
}

export interface Size {
  readonly width: number
  readonly height: number
}

export function newId(): string {
  return crypto.randomUUID()
}

export function isRedaction(markup: Markup): markup is RedactBox {
  return markup.kind === 'redact'
}

export function toPageRef(item: PageItem): PageRef {
  return {
    source: item.source,
    index: item.index,
    rotation: item.rotation,
    edits: item.markups.flatMap(placed => isRedaction(placed.markup) ? [] : [placed.markup]),
  }
}

/** Swaps width and height for quarter turns. */
export function rotatedSize(size: Size, rotation: number): Size {
  return rotation % 180 === 0 ? size : { width: size.height, height: size.width }
}

type Turn = 1 | -1

/** Maps a point on a page of `size` (before turning) to the page turned a quarter clockwise (1) or counter-clockwise (-1). */
function turnPoint(point: PdfPoint, size: Size, turn: Turn): PdfPoint {
  return turn === 1
    ? { x: size.height - point.y, y: point.x }
    : { x: point.y, y: size.width - point.x }
}

interface Box {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

/** Boxes that are drawn axis-aligned in the new orientation: their corners map, so width and height swap. */
function turnBox<T extends Box>(box: T, size: Size, turn: Turn): T {
  const a = turnPoint({ x: box.x, y: box.y }, size, turn)
  const b = turnPoint({ x: box.x + box.width, y: box.y + box.height }, size, turn)
  return { ...box, x: Math.min(a.x, b.x), y: Math.min(a.y, b.y), width: box.height, height: box.width }
}

/** A box on a page turned `rotation` degrees clockwise; `size` is the page's displayed size before turning. */
export function rotateBox<T extends Box>(box: T, size: Size, rotation: PageRotation): T {
  let turned = box
  let current = size
  for (let quarter = 0; quarter < rotation / 90; quarter += 1) {
    turned = turnBox(turned, current, 1)
    current = rotatedSize(current, 90)
  }
  return turned
}

/**
 * Moves a markup onto the turned page. Text and images stay upright (the engine always draws them
 * upright to the viewer), so only their position moves: their centre maps to the new centre.
 */
function turnMarkup(markup: Markup, size: Size, turn: Turn): Markup {
  switch (markup.kind) {
    case 'rect':
    case 'redact':
      return turnBox(markup, size, turn)
    case 'image':
    case 'text': {
      const bounds = markupBounds(markup)
      const centre = turnPoint({ x: bounds.x + bounds.width / 2, y: bounds.y + bounds.height / 2 }, size, turn)
      return { ...markup, x: centre.x - bounds.width / 2, y: centre.y - bounds.height / 2 }
    }
    case 'ink':
      return { ...markup, points: markup.points.map(point => turnPoint(point, size, turn)) }
  }
}

/** Turns a page a quarter clockwise (1) or counter-clockwise (-1); `size` is its displayed size before turning. */
export function turnPage(item: PageItem, size: Size, turn: Turn): PageItem {
  const rotation = (((item.rotation + turn * 90) % 360) + 360) % 360
  return {
    ...item,
    rotation: rotation === 90 || rotation === 180 || rotation === 270 ? rotation : 0,
    markups: item.markups.map(placed => ({ ...placed, markup: turnMarkup(placed.markup, size, turn) })),
  }
}

export function translateMarkup(markup: Markup, dx: number, dy: number): Markup {
  if (markup.kind === 'ink') {
    return { ...markup, points: markup.points.map(point => ({ x: point.x + dx, y: point.y + dy })) }
  }
  return { ...markup, x: markup.x + dx, y: markup.y + dy }
}

/** Resizes from the bottom-right corner. Images keep their aspect ratio and stay at least 12pt wide; boxes at least 6pt. */
export function resizeMarkup(markup: Markup, width: number, height: number): Markup {
  if (markup.kind === 'image') {
    const nextWidth = Math.max(12, width)
    return { ...markup, width: nextWidth, height: nextWidth * (markup.height / markup.width) }
  }
  if (markup.kind === 'rect' || markup.kind === 'redact') {
    return { ...markup, width: Math.max(6, width), height: Math.max(6, height) }
  }
  return markup
}

/** Bounding box of a markup in page points; text uses an estimate from its lines and size. */
export function markupBounds(markup: Markup): Box {
  switch (markup.kind) {
    case 'ink': {
      const xs = markup.points.map(point => point.x)
      const ys = markup.points.map(point => point.y)
      const pad = markup.width / 2
      const x = Math.min(...xs) - pad
      const y = Math.min(...ys) - pad
      return { x, y, width: Math.max(...xs) + pad - x, height: Math.max(...ys) + pad - y }
    }
    case 'text': {
      const lines = markup.text.split('\n')
      const longest = Math.max(1, ...lines.map(line => line.length))
      return { x: markup.x, y: markup.y, width: longest * markup.size * 0.55, height: lines.length * markup.size * 1.2 }
    }
    default:
      return { x: markup.x, y: markup.y, width: markup.width, height: markup.height }
  }
}

/**
 * How many edits separate the current state from the last save, walking the undo history.
 * Both stacks end at the state nearest `present`: `past.at(-1)` is the next undo, `future.at(-1)` the next redo.
 */
export function changesSinceSave(past: readonly Snapshot[], present: Snapshot, future: readonly Snapshot[], saved: Snapshot): number {
  if (present === saved) return 0
  const behind = past.lastIndexOf(saved)
  if (behind >= 0) return past.length - behind
  const ahead = future.lastIndexOf(saved)
  return ahead >= 0 ? future.length - ahead : 1
}
