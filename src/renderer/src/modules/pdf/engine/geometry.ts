import type { PDFPage } from '@cantoo/pdf-lib'
import type { PageRotation, PdfPoint } from './types'

export interface Rect {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

/**
 * A page's visible box in PDF user space plus the clockwise rotation a viewer
 * applies to it. Everything the UI sends is relative to the box *after* that
 * rotation, which is what pdf.js draws for `getViewport({ scale: 1 })`.
 */
export interface PageGeometry extends Rect {
  readonly rotation: PageRotation
}

export function normalizeRotation(angle: number): PageRotation {
  const quarters = (((Math.round(angle / 90) % 4) + 4) % 4)
  if (quarters === 1) return 90
  if (quarters === 2) return 180
  if (quarters === 3) return 270
  return 0
}

export function displayedSize(geometry: PageGeometry): { readonly width: number; readonly height: number } {
  return geometry.rotation === 90 || geometry.rotation === 270
    ? { width: geometry.height, height: geometry.width }
    : { width: geometry.width, height: geometry.height }
}

/** Displayed top-left coordinates (y down) to PDF user space (y up). */
export function toUserSpace(geometry: PageGeometry, point: PdfPoint): PdfPoint {
  const { x, y, width, height, rotation } = geometry
  if (rotation === 90) return { x: x + point.y, y: y + point.x }
  if (rotation === 180) return { x: x + width - point.x, y: y + point.y }
  if (rotation === 270) return { x: x + width - point.y, y: y + height - point.x }
  return { x: x + point.x, y: y + height - point.y }
}

/** Inverse of `toUserSpace`. */
export function toDisplayed(geometry: PageGeometry, point: PdfPoint): PdfPoint {
  const { x, y, width, height, rotation } = geometry
  if (rotation === 90) return { x: point.y - y, y: point.x - x }
  if (rotation === 180) return { x: width - (point.x - x), y: point.y - y }
  if (rotation === 270) return { x: height - (point.y - y), y: width - (point.x - x) }
  return { x: point.x - x, y: height - (point.y - y) }
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

export function pageGeometry(page: PDFPage, extraRotation: PageRotation = 0): PageGeometry {
  const box = visibleBox(page.getCropBox(), page.getMediaBox())
  return { ...box, rotation: normalizeRotation(page.getRotation().angle + extraRotation) }
}
