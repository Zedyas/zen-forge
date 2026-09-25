import { PDFName, PDFNumber, type PDFPage } from '@cantoo/pdf-lib'
import type { PageRotation, PdfPoint } from './types'

export interface Rect {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

/**
 * A page's visible box in PDF user space, the clockwise rotation a viewer applies to it, and the
 * size of its user-space unit. Everything the UI sends is relative to the box *after* that
 * rotation, in points, which is what pdf.js draws for `getViewport({ scale: 1 })`.
 */
export interface PageGeometry extends Rect {
  readonly rotation: PageRotation
  /** Points per user-space unit, from the page's /UserUnit (PDF 1.6); pdf.js scales the page by it. */
  readonly userUnit: number
}

/** A /Rotate value as pdf.js applies it: one that is not a multiple of 90 counts as 0. */
export function normalizeRotation(angle: number): PageRotation {
  if (!Number.isFinite(angle) || angle % 90 !== 0) return 0
  const quarters = (((angle / 90) % 4) + 4) % 4
  if (quarters === 1) return 90
  if (quarters === 2) return 180
  if (quarters === 3) return 270
  return 0
}

export function displayedSize(geometry: PageGeometry): { readonly width: number; readonly height: number } {
  const width = geometry.width * geometry.userUnit
  const height = geometry.height * geometry.userUnit
  return geometry.rotation === 90 || geometry.rotation === 270 ? { width: height, height: width } : { width, height }
}

/** Displayed top-left coordinates (y down, in points) to PDF user space (y up, in user units). */
export function toUserSpace(geometry: PageGeometry, displayed: PdfPoint): PdfPoint {
  const { x, y, width, height, rotation, userUnit } = geometry
  const point = { x: displayed.x / userUnit, y: displayed.y / userUnit }
  if (rotation === 90) return { x: x + point.y, y: y + point.x }
  if (rotation === 180) return { x: x + width - point.x, y: y + point.y }
  if (rotation === 270) return { x: x + width - point.y, y: y + height - point.x }
  return { x: x + point.x, y: y + height - point.y }
}

/** Inverse of `toUserSpace`. */
export function toDisplayed(geometry: PageGeometry, point: PdfPoint): PdfPoint {
  const { x, y, width, height, rotation, userUnit: unit } = geometry
  if (rotation === 90) return { x: (point.y - y) * unit, y: (point.x - x) * unit }
  if (rotation === 180) return { x: (width - (point.x - x)) * unit, y: (point.y - y) * unit }
  if (rotation === 270) return { x: (height - (point.y - y)) * unit, y: (width - (point.x - x)) * unit }
  return { x: (point.x - x) * unit, y: (height - (point.y - y)) * unit }
}

/**
 * A displayed rectangle stays axis-aligned in user space because rotations are
 * always quarter turns, so mapping two opposite corners is enough.
 */
export function rectToUserSpace(geometry: PageGeometry, rect: Rect): Rect {
  const a = toUserSpace(geometry, { x: rect.x, y: rect.y })
  const b = toUserSpace(geometry, { x: rect.x + rect.width, y: rect.y + rect.height })
  return {
    x: Math.min(a.x, b.x),
    y: Math.min(a.y, b.y),
    width: Math.abs(b.x - a.x),
    height: Math.abs(b.y - a.y),
  }
}

/** Inverse of `rectToUserSpace`. */
export function rectToDisplayed(geometry: PageGeometry, rect: Rect): Rect {
  const a = toDisplayed(geometry, { x: rect.x, y: rect.y })
  const b = toDisplayed(geometry, { x: rect.x + rect.width, y: rect.y + rect.height })
  return {
    x: Math.min(a.x, b.x),
    y: Math.min(a.y, b.y),
    width: Math.abs(b.x - a.x),
    height: Math.abs(b.y - a.y),
  }
}

function normalizeRect(rect: Rect): Rect {
  return {
    x: Math.min(rect.x, rect.x + rect.width),
    y: Math.min(rect.y, rect.y + rect.height),
    width: Math.abs(rect.width),
    height: Math.abs(rect.height),
  }
}

/** Matches pdf.js, which clips the crop box to the media box before rendering. */
function visibleBox(crop: Rect, media: Rect): Rect {
  const c = normalizeRect(crop)
  const m = normalizeRect(media)
  const x = Math.max(c.x, m.x)
  const y = Math.max(c.y, m.y)
  const right = Math.min(c.x + c.width, m.x + m.width)
  const top = Math.min(c.y + c.height, m.y + m.height)
  return right > x && top > y ? { x, y, width: right - x, height: top - y } : m
}

/** Matches pdf.js, which reads /UserUnit from the page itself and ignores a value that is not a positive number. */
function userUnitOf(page: PDFPage): number {
  const value = page.node.lookup(PDFName.of('UserUnit'))
  return value instanceof PDFNumber && value.asNumber() > 0 ? value.asNumber() : 1
}

export function pageGeometry(page: PDFPage, extraRotation: PageRotation = 0): PageGeometry {
  const box = visibleBox(page.getCropBox(), page.getMediaBox())
  const rotation = normalizeRotation(normalizeRotation(page.getRotation().angle) + extraRotation)
  return { ...box, rotation, userUnit: userUnitOf(page) }
}

/** What pdf.js reports for a page (`PDFPageProxy` has these), as a `PageGeometry`. */
export function shownGeometry(page: { readonly view: readonly number[]; readonly rotate: number; readonly userUnit: number }, extraRotation: PageRotation = 0): PageGeometry {
  const [x1 = 0, y1 = 0, x2 = 0, y2 = 0] = page.view
  return { x: x1, y: y1, width: x2 - x1, height: y2 - y1, rotation: normalizeRotation(page.rotate + extraRotation), userUnit: page.userUnit }
}

/** Whether two geometries place the same displayed point at the same point of the page, to a hundredth of a unit. */
export function sameGeometry(a: PageGeometry, b: PageGeometry): boolean {
  const close = (p: number, q: number): boolean => Math.abs(p - q) < 0.01
  return a.rotation === b.rotation && close(a.userUnit, b.userUnit)
    && close(a.x, b.x) && close(a.y, b.y) && close(a.width, b.width) && close(a.height, b.height)
}
