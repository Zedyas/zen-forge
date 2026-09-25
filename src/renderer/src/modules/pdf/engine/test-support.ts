import { decodePDFRawStream, PDFDocument, PDFName, PDFRawStream, PDFString, StandardFonts, type PDFPage, type PDFRef } from '@cantoo/pdf-lib'
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs'
import { shownGeometry, type PageGeometry } from './geometry'

interface Resolvers<T> {
  promise: Promise<T>
  resolve: (value: T | PromiseLike<T>) => void
  reject: (reason?: unknown) => void
}

// pdf.js 6 calls Promise.withResolvers, which Node 20 does not have.
const promiseConstructor: PromiseConstructor & { withResolvers?: <T>() => Resolvers<T> } = Promise
if (promiseConstructor.withResolvers === undefined) {
  promiseConstructor.withResolvers = <T>(): Resolvers<T> => {
    let resolve: (value: T | PromiseLike<T>) => void = () => undefined
    let reject: (reason?: unknown) => void = () => undefined
    const promise = new Promise<T>((res, rej) => {
      resolve = res
      reject = rej
    })
    return { promise, resolve, reject }
  }
}

export interface PlacedText {
  readonly text: string
  /** Baseline origin in displayed top-left coordinates, straight from pdf.js. */
  readonly x: number
  readonly y: number
}

async function withDocument<T>(
  bytes: Uint8Array,
  read: (doc: pdfjs.PDFDocumentProxy) => Promise<T>,
): Promise<T> {
  const task = pdfjs.getDocument({ data: new Uint8Array(bytes), useSystemFonts: false })
  try {
    return await read(await task.promise)
  } finally {
    await task.destroy()
  }
}

/** The text pdf.js can select on each page, in reading order. */
export async function extractPageTexts(bytes: Uint8Array): Promise<string[]> {
  return withDocument(bytes, async doc => {
    const texts: string[] = []
    for (let number = 1; number <= doc.numPages; number += 1) {
      const page = await doc.getPage(number)
      const content = await page.getTextContent()
      texts.push(content.items.map(item => ('str' in item ? item.str : '')).join(''))
    }
    return texts
  })
}

/** Each page as pdf.js shows it, which is the page the editor draws redaction boxes on. */
export async function shownPages(bytes: Uint8Array): Promise<PageGeometry[]> {
  return withDocument(bytes, async doc => {
    const pages: PageGeometry[] = []
    for (let number = 1; number <= doc.numPages; number += 1) pages.push(shownGeometry(await doc.getPage(number)))
    return pages
  })
}

export async function extractPlacedText(bytes: Uint8Array, pageIndex: number): Promise<PlacedText[]> {
  return withDocument(bytes, async doc => {
    const page = await doc.getPage(pageIndex + 1)
    const viewport = page.getViewport({ scale: 1 })
    const content = await page.getTextContent()
    return content.items.flatMap(item => {
      if (!('str' in item) || item.str.length === 0) return []
      // applyTransform rewrites the point in place; the viewport transform is
      // what pdf.js itself uses to turn user space into on-screen coordinates.
      const point = [item.transform[4], item.transform[5]]
      pdfjs.Util.applyTransform(point, viewport.transform)
      return [{ text: item.str, x: point[0], y: point[1] }]
    })
  })
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let n = 0; n < 256; n += 1) {
    let value = n
    for (let bit = 0; bit < 8; bit += 1) {
      value = (value & 1) === 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1
    }
    table[n] = value >>> 0
  }
  return table
})()

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff
  for (const byte of bytes) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8)
  return (crc ^ 0xffffffff) >>> 0
}

function adler32(bytes: Uint8Array): number {
  let low = 1
  let high = 0
  for (const byte of bytes) {
    low = (low + byte) % 65521
    high = (high + low) % 65521
  }
  return (((high << 16) | low) >>> 0)
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const total = parts.reduce((sum, part) => sum + part.length, 0)
  const out = new Uint8Array(total)
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length)
  const view = new DataView(out.buffer)
  view.setUint32(0, data.length)
  for (let i = 0; i < 4; i += 1) out[4 + i] = type.charCodeAt(i)
  out.set(data, 8)
  view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)))
  return out
}

/** Zlib stream made only of stored (uncompressed) deflate blocks. */
function storedZlib(raw: Uint8Array): Uint8Array {
  const parts: Uint8Array[] = [new Uint8Array([0x78, 0x01])]
  const blockSize = 65535
  for (let offset = 0; offset < raw.length; offset += blockSize) {
    const slice = raw.subarray(offset, offset + blockSize)
    const isFinal = offset + blockSize >= raw.length
    const header = new Uint8Array(5)
    header[0] = isFinal ? 1 : 0
    header[1] = slice.length & 0xff
    header[2] = (slice.length >>> 8) & 0xff
    header[3] = ~slice.length & 0xff
    header[4] = (~slice.length >>> 8) & 0xff
    parts.push(header, slice)
  }
  const trailer = new Uint8Array(4)
  new DataView(trailer.buffer).setUint32(0, adler32(raw))
  parts.push(trailer)
  return concat(parts)
}

/** A solid-colour 8-bit RGB PNG, optionally carrying a pHYs density. */
export function encodePng(
  width: number,
  height: number,
  colour: readonly [number, number, number],
  dpi?: number,
): Uint8Array {
  const raw = new Uint8Array((1 + width * 3) * height)
  for (let row = 0; row < height; row += 1) {
    const start = row * (1 + width * 3)
    for (let column = 0; column < width; column += 1) {
      raw[start + 1 + column * 3] = colour[0]
      raw[start + 2 + column * 3] = colour[1]
      raw[start + 3 + column * 3] = colour[2]
    }
  }

  const header = new Uint8Array(13)
  const headerView = new DataView(header.buffer)
  headerView.setUint32(0, width)
  headerView.setUint32(4, height)
  header[8] = 8
  header[9] = 2

  const chunks = [chunk('IHDR', header)]
  if (dpi !== undefined) {
    const phys = new Uint8Array(9)
    const physView = new DataView(phys.buffer)
    const perMetre = Math.round(dpi / 0.0254)
    physView.setUint32(0, perMetre)
    physView.setUint32(4, perMetre)
    phys[8] = 1
    chunks.push(chunk('pHYs', phys))
  }
  chunks.push(chunk('IDAT', storedZlib(raw)), chunk('IEND', new Uint8Array(0)))

  return concat([new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]), ...chunks])
}

/**
 * Whether `text` survives anywhere in the file, not only on its pages: every stream is decoded
 * and searched for the text and for the hex form pdf-lib writes standard-font strings in.
 */
export async function fileContainsText(bytes: Uint8Array, text: string): Promise<boolean> {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false })
  const hex = Array.from(text, character => character.charCodeAt(0).toString(16).padStart(2, '0')).join('')
  const needles = [text, hex.toUpperCase(), hex.toLowerCase()]
  return doc.context.enumerateIndirectObjects().some(([, object]) => {
    let content = object.toString()
    if (object instanceof PDFRawStream) {
      let decoded: Uint8Array
      try {
        decoded = decodePDFRawStream(object).decode()
      } catch {
        decoded = object.contents
      }
      content += new TextDecoder('latin1').decode(decoded)
    }
    return needles.some(needle => content.includes(needle))
  })
}

export interface LayeredPdf {
  readonly doc: PDFDocument
  readonly page: PDFPage
  readonly layers: { readonly hidden: PDFRef; readonly shown: PDFRef }
}

/**
 * One 400 × 600 page drawn by `content`, with two layers: `hidden` is OFF in the default
 * configuration and `shown` is ON. Content names them as /hidden and /shown (for `/OC … BDC`),
 * Helvetica as /F1, and each of `forms` as a form XObject, optionally in a layer through /OC.
 */
export async function layeredPdf(
  content: string,
  forms: Readonly<Record<string, { readonly content: string; readonly layer?: 'hidden' | 'shown' }>> = {},
): Promise<LayeredPdf> {
  const doc = await PDFDocument.create()
  const { context } = doc
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const page = doc.addPage([400, 600])
  const layers = {
    hidden: context.register(context.obj({ Type: 'OCG', Name: PDFString.of('Hidden layer') })),
    shown: context.register(context.obj({ Type: 'OCG', Name: PDFString.of('Shown layer') })),
  }
  doc.catalog.set(PDFName.of('OCProperties'), context.obj({ OCGs: [layers.hidden, layers.shown], D: { OFF: [layers.hidden] } }))
  const xobjects = Object.fromEntries(Object.entries(forms).map(([name, form]) => [name, context.register(context.stream(form.content, {
    Type: 'XObject',
    Subtype: 'Form',
    BBox: [0, 0, 400, 600],
    Resources: { Font: { F1: font.ref } },
    ...(form.layer === undefined ? {} : { OC: layers[form.layer] }),
  }))]))
  page.node.set(PDFName.of('Resources'), context.obj({ Font: { F1: font.ref }, Properties: layers, XObject: xobjects }))
  page.node.set(PDFName.of('Contents'), context.register(context.stream(content)))
  return { doc, page, layers }
}
