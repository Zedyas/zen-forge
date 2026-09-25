import { PDFDocument, degrees } from '@cantoo/pdf-lib'
import { describe, expect, it } from 'vitest'
import {
  displayedSize,
  normalizeRotation,
  pageGeometry,
  rectToUserSpace,
  toDisplayed,
  toUserSpace,
  type PageGeometry,
} from './geometry'

/** A crop box that does not start at the origin, so offsets cannot hide behind zeros. */
const CROP = { x: 20, y: 30, width: 200, height: 400 }

const geometryAt = (rotation: PageGeometry['rotation']): PageGeometry => ({ ...CROP, rotation, userUnit: 1 })

describe('displayed to user space conversion', () => {
  it('swaps the displayed size on quarter turns', () => {
    expect(displayedSize(geometryAt(0))).toEqual({ width: 200, height: 400 })
    expect(displayedSize(geometryAt(90))).toEqual({ width: 400, height: 200 })
    expect(displayedSize(geometryAt(180))).toEqual({ width: 200, height: 400 })
    expect(displayedSize(geometryAt(270))).toEqual({ width: 400, height: 200 })
  })

  it('maps the displayed top-left corner onto the right user-space corner', () => {
    expect(toUserSpace(geometryAt(0), { x: 0, y: 0 })).toEqual({ x: 20, y: 430 })
    expect(toUserSpace(geometryAt(90), { x: 0, y: 0 })).toEqual({ x: 20, y: 30 })
    expect(toUserSpace(geometryAt(180), { x: 0, y: 0 })).toEqual({ x: 220, y: 30 })
    expect(toUserSpace(geometryAt(270), { x: 0, y: 0 })).toEqual({ x: 220, y: 430 })
  })

  it('maps an interior point for every rotation', () => {
    const point = { x: 10, y: 15 }
    expect(toUserSpace(geometryAt(0), point)).toEqual({ x: 30, y: 415 })
    expect(toUserSpace(geometryAt(90), point)).toEqual({ x: 35, y: 40 })
    expect(toUserSpace(geometryAt(180), point)).toEqual({ x: 210, y: 45 })
    expect(toUserSpace(geometryAt(270), point)).toEqual({ x: 205, y: 420 })
  })

  it('keeps the displayed bottom-right corner inside the page for every rotation', () => {
    for (const rotation of [0, 90, 180, 270] as const) {
      const geometry = geometryAt(rotation)
      const size = displayedSize(geometry)
      const corner = toUserSpace(geometry, { x: size.width, y: size.height })
      expect(corner.x).toBeGreaterThanOrEqual(CROP.x)
      expect(corner.x).toBeLessThanOrEqual(CROP.x + CROP.width)
      expect(corner.y).toBeGreaterThanOrEqual(CROP.y)
      expect(corner.y).toBeLessThanOrEqual(CROP.y + CROP.height)
    }
  })

  it('round-trips through toDisplayed for every rotation', () => {
    const point = { x: 37, y: 91 }
    for (const rotation of [0, 90, 180, 270] as const) {
      const geometry = geometryAt(rotation)
      expect(toDisplayed(geometry, toUserSpace(geometry, point))).toEqual(point)
    }
  })

  it('keeps a displayed rectangle axis-aligned and the right way up', () => {
    const rect = { x: 10, y: 20, width: 30, height: 40 }
    expect(rectToUserSpace(geometryAt(0), rect)).toEqual({ x: 30, y: 370, width: 30, height: 40 })
    expect(rectToUserSpace(geometryAt(90), rect)).toEqual({ x: 40, y: 40, width: 40, height: 30 })
    expect(rectToUserSpace(geometryAt(180), rect)).toEqual({ x: 180, y: 50, width: 30, height: 40 })
    expect(rectToUserSpace(geometryAt(270), rect)).toEqual({ x: 160, y: 390, width: 40, height: 30 })
  })

  it('normalizes odd and negative /Rotate values, and ignores ones that are not quarter turns as pdf.js does', () => {
    expect(normalizeRotation(-90)).toBe(270)
    expect(normalizeRotation(450)).toBe(90)
    expect(normalizeRotation(0)).toBe(0)
    expect(normalizeRotation(45)).toBe(0)
    expect(normalizeRotation(135)).toBe(0)
  })
})

describe('pageGeometry', () => {
  it('reads the crop box and adds the user rotation to the page rotation', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 500])
    page.setCropBox(CROP.x, CROP.y, CROP.width, CROP.height)
    page.setRotation(degrees(90))

    expect(pageGeometry(page)).toEqual({ ...CROP, rotation: 90, userUnit: 1 })
    expect(pageGeometry(page, 180)).toEqual({ ...CROP, rotation: 270, userUnit: 1 })
    expect(pageGeometry(page, 270)).toEqual({ ...CROP, rotation: 0, userUnit: 1 })
  })
})
