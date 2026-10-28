import { PDFDocument } from '@cantoo/pdf-lib'
import { letterPaper } from '@shared/print'
import { describe, expect, it } from 'vitest'
import { imagesToPdf, jpegOrientation, orientedPlacement } from './images'
import { encodePng } from './test-support'

/** A JPEG header pdf-lib can embed: an Exif block holding `orientation`, then a `width` × `height` frame. */
function jpeg(width: number, height: number, orientation?: number, littleEndian = false): Uint8Array {
  const bytes: number[] = [0xff, 0xd8]
  const short = (value: number, little = false): number[] => little ? [value & 0xff, value >> 8] : [value >> 8, value & 0xff]
  const long = (value: number, little: boolean): number[] => little ? [...short(value & 0xffff, true), 0, 0] : [0, 0, ...short(value)]
  if (orientation !== undefined) {
    const tiff = [
      ...(littleEndian ? [0x49, 0x49] : [0x4d, 0x4d]), ...short(42, littleEndian), ...long(8, littleEndian),
      ...short(1, littleEndian),
      ...short(0x0112, littleEndian), ...short(3, littleEndian), ...long(1, littleEndian), ...short(orientation, littleEndian), 0, 0,
      ...long(0, littleEndian),
    ]
    bytes.push(0xff, 0xe1, ...short(2 + 6 + tiff.length), ...[...'Exif'].map(char => char.charCodeAt(0)), 0, 0, ...tiff)
  }
  bytes.push(0xff, 0xc0, 0, 17, 8, ...short(height), ...short(width), 3, 1, 0x11, 0, 2, 0x11, 1, 3, 0x11, 1)
  bytes.push(0xff, 0xd9)
  return Uint8Array.from(bytes)
}

/** Where a point of the stored image (u right, v up, both 0–1) lands on the page. */
function place(matrix: readonly number[], u: number, v: number): [number, number] {
  const [a = 0, b = 0, c = 0, d = 0, e = 0, f = 0] = matrix
  return [a * u + c * v + e, b * u + d * v + f]
}

describe('imagesToPdf', () => {
  it('sizes each page from the image density rather than its pixel count', async () => {
    const output = await imagesToPdf([
      { bytes: encodePng(200, 100, [255, 0, 0], 144), type: 'png' },
      { bytes: encodePng(200, 100, [0, 255, 0]), type: 'png' },
    ], letterPaper)

    const sizes = (await PDFDocument.load(output)).getPages().map(page => page.getSize())
    expect(sizes[0].width).toBeCloseTo(100, 1)
    expect(sizes[0].height).toBeCloseTo(50, 1)
    expect(sizes[1]).toEqual({ width: 200, height: 100 })
  })

  it('reads the EXIF orientation in either byte order', () => {
    expect(jpegOrientation(jpeg(4, 3, 6))).toBe(6)
    expect(jpegOrientation(jpeg(4, 3, 8, true))).toBe(8)
    expect(jpegOrientation(jpeg(4, 3))).toBe(1)
  })

  it('draws the stored top-left corner where the orientation puts it', () => {
    // The stored image's top-left corner is (0, 1) in the unit square; the page is 300 × 400.
    expect(place(orientedPlacement(1, 300, 400), 0, 1)).toEqual([0, 400])
    expect(place(orientedPlacement(3, 300, 400), 0, 1)).toEqual([300, 0])
    // A phone held upright stores its picture turned left (6): its top-left belongs at the top right.
    expect(place(orientedPlacement(6, 300, 400), 0, 1)).toEqual([300, 400])
    expect(place(orientedPlacement(6, 300, 400), 1, 0)).toEqual([0, 0])
    expect(place(orientedPlacement(8, 300, 400), 0, 1)).toEqual([0, 0])
    expect(place(orientedPlacement(5, 300, 400), 1, 1)).toEqual([0, 0])
  })

  it('turns a sideways phone photo upright and fits it to the paper', async () => {
    const output = await imagesToPdf([{ bytes: jpeg(4032, 3024, 6), type: 'jpeg' }], letterPaper)

    // Upright it is 3024 × 4032 points at 72 dpi; on Letter it is as tall as the paper.
    const { width, height } = (await PDFDocument.load(output)).getPage(0).getSize()
    expect(width).toBeCloseTo(594, 0)
    expect(height).toBeCloseTo(792, 0)
  })
})
