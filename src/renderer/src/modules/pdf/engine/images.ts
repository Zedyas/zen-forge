import { PDFDocument, type PDFImage } from '@cantoo/pdf-lib'
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

function intrinsicSize(image: ImageSource, embedded: PDFImage): { width: number; height: number } {
  const density = image.type === 'png' ? pngDensity(image.bytes) : jpegDensity(image.bytes)
  const dpiX = density?.x ?? DEFAULT_DPI
  const dpiY = density?.y ?? DEFAULT_DPI
  return { width: (embedded.width * 72) / dpiX, height: (embedded.height * 72) / dpiY }
}

/** One page per image, each page the image's own size in points (from its stored density). */
export async function imagesToPdf(images: readonly ImageSource[]): Promise<Uint8Array> {
  if (images.length === 0) throw new Error('imagesToPdf needs at least one image')
  const doc = await PDFDocument.create()

  for (const image of images) {
    const embedded = image.type === 'png'
      ? await doc.embedPng(image.bytes)
      : await doc.embedJpg(image.bytes)
    const { width, height } = intrinsicSize(image, embedded)
    doc.addPage([width, height]).drawImage(embedded, { x: 0, y: 0, width, height })
  }

  return doc.save()
}
