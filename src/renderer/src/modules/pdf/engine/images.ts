import { concatTransformationMatrix, drawObject, PDFDocument, popGraphicsState, pushGraphicsState, type PDFImage } from '@cantoo/pdf-lib'
import type { Paper } from '@shared/print'
import type { ImageSource } from './types'

const DEFAULT_DPI = 72

interface Density {
  readonly x: number
  readonly y: number
}

/** Reads a PNG `pHYs` chunk, which stores pixels per unit and a unit flag (1 = metre). */
function pngDensity(bytes: Uint8Array): Density | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let offset = 8
  while (offset + 8 <= bytes.length) {
    const length = view.getUint32(offset)
    const type = String.fromCharCode(...bytes.subarray(offset + 4, offset + 8))
    if (type === 'pHYs' && length >= 9) {
      const unit = bytes[offset + 8 + 8]
      if (unit !== 1) return undefined
      const perMetreX = view.getUint32(offset + 8)
      const perMetreY = view.getUint32(offset + 12)
      if (perMetreX === 0 || perMetreY === 0) return undefined
      return { x: perMetreX * 0.0254, y: perMetreY * 0.0254 }
    }
    if (type === 'IDAT' || type === 'IEND') return undefined
    offset += 12 + length
  }
  return undefined
}

/** Reads the JFIF APP0 segment, the only density marker a JPEG carries cheaply. */
function jpegDensity(bytes: Uint8Array): Density | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let offset = 2
  while (offset + 4 <= bytes.length && bytes[offset] === 0xff) {
    const marker = bytes[offset + 1]
    const length = view.getUint16(offset + 2)
    if (marker === 0xe0 && length >= 16) {
      const units = bytes[offset + 11]
      const densityX = view.getUint16(offset + 12)
      const densityY = view.getUint16(offset + 14)
      if (densityX === 0 || densityY === 0) return undefined
      if (units === 1) return { x: densityX, y: densityY }
      if (units === 2) return { x: densityX * 2.54, y: densityY * 2.54 }
      return undefined
    }
    if (marker === 0xda) return undefined
    offset += 2 + length
  }
  return undefined
}

/**
 * The EXIF orientation (1–8) a camera stores in a JPEG's APP1 segment instead of turning the
 * pixels; 1, upright, when there is none. A damaged Exif block also reads as upright.
 */
export function jpegOrientation(bytes: Uint8Array): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  try {
    let offset = 2
    while (offset + 4 <= bytes.length && bytes[offset] === 0xff) {
      const marker = bytes[offset + 1]
      const length = view.getUint16(offset + 2)
      if (marker === 0xe1 && String.fromCharCode(...bytes.subarray(offset + 4, offset + 10)) === 'Exif\0\0') {
        // A TIFF header follows: byte order (II little-endian, MM big-endian), 42, then the first directory's offset.
        const tiff = offset + 10
        const little = bytes[tiff] === 0x49
        const directory = tiff + view.getUint32(tiff + 4, little)
        const entries = view.getUint16(directory, little)
        for (let entry = 0; entry < entries; entry += 1) {
          const at = directory + 2 + entry * 12
          if (view.getUint16(at, little) !== 0x0112) continue
          const orientation = view.getUint16(at + 8, little)
          return orientation >= 1 && orientation <= 8 ? orientation : 1
        }
        return 1
      }
      if (marker === 0xda) return 1
      offset += 2 + length
    }
  } catch {
    // Read past the end of the file.
  }
  return 1
}

type Matrix = readonly [number, number, number, number, number, number]

/**
 * The matrix that draws an image (the unit square, its stored pixels) upright on a `width` ×
 * `height` page, for an EXIF orientation: 2–4 mirror or half-turn it, 5–8 also turn it a quarter,
 * so for those the page is as wide as the stored image is tall.
 */
export function orientedPlacement(orientation: number, width: number, height: number): Matrix {
  switch (orientation) {
    case 2: return [-width, 0, 0, height, width, 0]
    case 3: return [-width, 0, 0, -height, width, height]
    case 4: return [width, 0, 0, -height, 0, height]
    case 5: return [0, -height, -width, 0, width, height]
    case 6: return [0, -height, width, 0, 0, height]
    case 7: return [0, height, width, 0, 0, 0]
    case 8: return [0, height, -width, 0, width, 0]
    default: return [width, 0, 0, height, 0, 0]
  }
}

function intrinsicSize(image: ImageSource, embedded: PDFImage): { width: number; height: number } {
  const density = image.type === 'png' ? pngDensity(image.bytes) : jpegDensity(image.bytes)
  const dpiX = density?.x ?? DEFAULT_DPI
  const dpiY = density?.y ?? DEFAULT_DPI
  return { width: (embedded.width * 72) / dpiX, height: (embedded.height * 72) / dpiY }
}

/** Shrinks a page, keeping its shape, to fit `paper` turned the same way; a page that fits keeps its size. */
function fitToPaper(size: { width: number; height: number }, paper: Paper): { width: number; height: number } {
  const long = (paper.height * 72) / 25.4
  const short = (paper.width * 72) / 25.4
  const [maxWidth, maxHeight] = size.width > size.height ? [long, short] : [short, long]
  const scale = Math.min(1, maxWidth / size.width, maxHeight / size.height)
  return { width: size.width * scale, height: size.height * scale }
}

/**
 * One page per image, each page the image's own size in points (from its stored density), upright
 * as the camera's EXIF orientation says, and shrunk to fit `paper` when larger. A photo stored at
 * 72 dpi would otherwise make a page over a metre wide.
 */
export async function imagesToPdf(images: readonly ImageSource[], paper: Paper): Promise<Uint8Array> {
  if (images.length === 0) throw new Error('imagesToPdf needs at least one image')
  const doc = await PDFDocument.create()

  for (const image of images) {
    const embedded = image.type === 'png'
      ? await doc.embedPng(image.bytes)
      : await doc.embedJpg(image.bytes)
    const orientation = image.type === 'jpeg' ? jpegOrientation(image.bytes) : 1
    const stored = intrinsicSize(image, embedded)
    const turned = orientation >= 5 ? { width: stored.height, height: stored.width } : stored
    const { width, height } = fitToPaper(turned, paper)
    const page = doc.addPage([width, height])
    page.pushOperators(
      pushGraphicsState(),
      concatTransformationMatrix(...orientedPlacement(orientation, width, height)),
      drawObject(page.node.newXObject('Image', embedded.ref)),
      popGraphicsState(),
    )
  }

  return doc.save()
}
