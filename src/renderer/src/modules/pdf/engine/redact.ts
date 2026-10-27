import {
  decodePDFRawStream,
  PDFArray,
  PDFDict,
  PDFName,
  PDFNumber,
  PDFObjectParser,
  PDFRawStream,
  PDFRef,
  type PDFDocument,
  type PDFObject,
  type PDFPage,
} from '@cantoo/pdf-lib'
import { zlibSync } from 'fflate'
import { scanContent } from './content-stream'
import { removeWidgets } from './forms'
import { pageGeometry, rectToUserSpace, sameGeometry, type Rect } from './geometry'
import { loadDocument } from './inspect'
import { applyEdits, removeHiddenLayers, type Edit } from './layers'
import { dropUnreachable } from './prune'
import type { RedactionRequest, RedactOptions } from './types'

/*
 * Redaction the way Acrobat Pro applies it: only what lies under a box is removed (text, the
 * covered pixels of images, line art the box covers) and the rest of the page keeps its real
 * text. Three passes:
 *   1. pdf-lib marks each box as a standard /Redact annotation in page space, and removes what
 *      would otherwise survive beside the content: form fields, links and comments touching a
 *      box, and the page's stored thumbnail.
 *   2. MuPDF applies the /Redact annotations, rewriting the page content itself.
 *   3. pdf-lib removes text copies outside the page content (tag alt text on redacted pages,
 *      and on request the document's hidden information), then deletes every unreferenced object.
 */

function numbers(value: PDFObject | undefined): number[] {
  return value instanceof PDFArray ? value.asArray().map(item => item instanceof PDFNumber ? item.asNumber() : 0) : []
}

/** An annotation's /Rect as a normalized rectangle in user space. */
function annotationRect(annotation: PDFDict): Rect | undefined {
  const [x1 = 0, y1 = 0, x2 = 0, y2 = 0] = numbers(annotation.get(PDFName.of('Rect')))
  return { x: Math.min(x1, x2), y: Math.min(y1, y2), width: Math.abs(x2 - x1), height: Math.abs(y2 - y1) }
}

function overlaps(a: Rect, b: Rect): boolean {
  return a.x < b.x + b.width && b.x < a.x + a.width && a.y < b.y + b.height && b.y < a.y + a.height
}

function touchesAny(rect: Rect | undefined, areas: readonly Rect[]): boolean {
  return rect !== undefined && areas.some(area => overlaps(rect, area))
}

function annotationRefs(page: PDFPage): PDFRef[] {
  const annots = page.node.Annots()
  return annots === undefined ? [] : annots.asArray().filter((item): item is PDFRef => item instanceof PDFRef)
}

/** Removes links, comments and other annotations that touch a box, with their pop-up notes. */
function removeTouchedAnnotations(doc: PDFDocument, page: PDFPage, areas: readonly Rect[]): void {
  for (const ref of annotationRefs(page)) {
    const annotation = doc.context.lookup(ref)
    if (!(annotation instanceof PDFDict) || !touchesAny(annotationRect(annotation), areas)) continue
    page.node.removeAnnot(ref)
    const popup = annotation.get(PDFName.of('Popup'))
    if (popup instanceof PDFRef) page.node.removeAnnot(popup)
  }
}

async function markRedactions(bytes: Uint8Array, requests: readonly RedactionRequest[]): Promise<Uint8Array> {
  const doc = await loadDocument(bytes)
  const pages = doc.getPages()
  const areas = new Map<PDFPage, readonly Rect[]>()
  for (const { index, shown, boxes } of requests) {
    const page = pages[index]
    if (page === undefined) throw new Error(`No page at index ${index}`)
    const geometry = pageGeometry(page)
    // The boxes were drawn on the page as shown. A page that differs in size, turn or scale would
    // take them somewhere else, and leave uncovered what they were drawn over.
    if (!sameGeometry(geometry, shown)) {
      throw new Error(`Zen Suzu can’t redact page ${index + 1} safely: the saved page would not match the page shown, so the boxes could miss. Nothing was saved. Redact this PDF in Adobe Acrobat instead.`)
    }
    areas.set(page, boxes.map(box => rectToUserSpace(geometry, box)))
  }

  // Form fields (value included) whose widgets all touch a box, and the touching widgets of the rest.
  removeWidgets(doc, (widget, page) => page !== undefined && touchesAny(widget.getRectangle(), areas.get(page) ?? []))
  for (const [page, rects] of areas) {
    removeTouchedAnnotations(doc, page, rects)
    // A stored thumbnail, application data (/PieceInfo) and page metadata can each hold a copy of the page.
    for (const key of ['Thumb', 'PieceInfo', 'Metadata']) page.node.delete(PDFName.of(key))
    for (const rect of rects) {
      const annotation = doc.context.obj({
        Type: 'Annot',
        Subtype: 'Redact',
        Rect: [rect.x, rect.y, rect.x + rect.width, rect.y + rect.height],
      })
      page.node.addAnnot(doc.context.register(annotation))
    }
  }
  return doc.save({ updateFieldAppearances: false })
}

/** MuPDF rewrites each page's content: covered text removed, covered pixels blanked, covered line art dropped. */
async function applyRedactions(bytes: Uint8Array, indexes: readonly number[]): Promise<Uint8Array> {
  // Loaded on first use: the WebAssembly module is large and only redaction needs it.
  const mupdf = await import('mupdf')
  const document = new mupdf.PDFDocument(bytes)
  try {
    for (const index of indexes) {
      const page = document.loadPage(index)
      page.applyRedactions(
        true,
        mupdf.PDFPage.REDACT_IMAGE_PIXELS,
        mupdf.PDFPage.REDACT_LINE_ART_REMOVE_IF_COVERED,
        mupdf.PDFPage.REDACT_TEXT_REMOVE,
      )
      page.destroy()
    }
    return document.saveToBuffer('garbage=compact').asUint8Array().slice()
  } finally {
    document.destroy()
  }
}

/**
 * Accessibility tags can quote a page's content as alt text or replacement text. On redacted
 * pages that text may describe what was removed, so it goes; the tag structure stays.
 */
function removeTagText(doc: PDFDocument, redactedPages: ReadonlySet<string>): void {
  const root = doc.catalog.lookupMaybe(PDFName.of('StructTreeRoot'), PDFDict)
  if (root === undefined) return
  const visited = new Set<PDFObject>()
  const visit = (node: PDFObject | undefined, inheritedPage: string | undefined): void => {
    const object = node instanceof PDFRef ? doc.context.lookup(node) : node
    if (object === undefined || visited.has(object)) return
    visited.add(object)
    if (object instanceof PDFArray) {
      object.asArray().forEach(child => visit(child, inheritedPage))
      return
    }
    if (!(object instanceof PDFDict)) return
    const pageOf = (dict: PDFObject | undefined): string | undefined => {
      const pg = dict instanceof PDFDict ? dict.get(PDFName.of('Pg')) : undefined
      return pg instanceof PDFRef ? pg.tag : undefined
    }
    const page = pageOf(object) ?? inheritedPage
    // An element's content can also name its page on each marked-content or object reference in /K.
    const kids = doc.context.lookup(object.get(PDFName.of('K')))
    const pages = [page, ...(kids instanceof PDFArray ? kids.asArray() : [kids]).map(pageOf)]
    if (pages.some(tag => tag !== undefined && redactedPages.has(tag))) {
      for (const key of ['Alt', 'ActualText', 'E', 'T']) object.delete(PDFName.of(key))
    }
    visit(object.get(PDFName.of('K')), page)
  }
  visit(root.get(PDFName.of('K')), undefined)
}

/** Property-list entries that can quote what marked content marks: replacement text, alternate description and expansion. */
const quotingKeys = ['ActualText', 'Alt', 'E'].map(key => PDFName.of(key))

/** Deletes the quoting entries of a property list; says whether it had any. */
function dropQuotes(properties: PDFDict): boolean {
  const quoted = quotingKeys.filter(key => properties.has(key))
  quoted.forEach(key => properties.delete(key))
  return quoted.length > 0
}

/** A content stream's property lists written inline (`/Span <</ActualText (…)>> BDC`) without their quoting entries, or undefined when none had any. */
function contentWithoutQuotes(doc: PDFDocument, content: Uint8Array): Uint8Array | undefined {
  // Without `<<` there is no inline property list, and nothing to scan for.
  if (!new TextDecoder('latin1').decode(content).includes('<<')) return undefined
  const edits: Edit[] = scanContent(content).flatMap(({ operator, operands: [tag, list], start, end }) => {
    if ((operator !== 'BDC' && operator !== 'DP') || tag?.startsWith('/') !== true || list?.startsWith('<<') !== true) return []
    const properties = PDFObjectParser.forBytes(Uint8Array.from(list, character => character.charCodeAt(0)), doc.context).parseObject()
    if (!(properties instanceof PDFDict) || !dropQuotes(properties)) return []
    return [{ start, end, replacement: `${PDFName.of(tag.slice(1)).toString()} ${properties.toString()} ${operator}` }]
  })
  return edits.length === 0 ? undefined : applyEdits(content, edits)
}

function withoutQuotes(doc: PDFDocument, stream: PDFObject | undefined): void {
  if (!(stream instanceof PDFRawStream)) return
  const cleaned = contentWithoutQuotes(doc, decodePDFRawStream(stream).decode())
  if (cleaned === undefined) return
  stream.dict.delete(PDFName.of('DecodeParms'))
  stream.dict.set(PDFName.of('Filter'), PDFName.of('FlateDecode'))
  stream.updateContents(zlibSync(cleaned))
}

/**
 * Marked content can quote the text it marks, as /ActualText or /Alt in its property list, and
 * MuPDF keeps those lists when it removes the text. On a redacted page they lose these entries:
 * inline in the content of the page and of the forms it draws, and named in their resources.
 */
function removeMarkedText(doc: PDFDocument, page: PDFPage): void {
  const visited = new Set<string>()
  const clean = (resources: PDFDict | undefined): void => {
    for (const value of resources?.lookupMaybe(PDFName.of('Properties'), PDFDict)?.values() ?? []) {
      const properties = doc.context.lookup(value)
      if (properties instanceof PDFDict) dropQuotes(properties)
    }
    for (const value of resources?.lookupMaybe(PDFName.of('XObject'), PDFDict)?.values() ?? []) {
      const form = doc.context.lookup(value)
      if (!(value instanceof PDFRef) || visited.has(value.tag) || !(form instanceof PDFRawStream)) continue
      visited.add(value.tag)
      if (form.dict.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText() !== 'Form') continue
      withoutQuotes(doc, form)
      clean(form.dict.lookupMaybe(PDFName.of('Resources'), PDFDict) ?? resources)
    }
  }
  const contents = doc.context.lookup(page.node.get(PDFName.of('Contents')))
  for (const stream of contents instanceof PDFArray ? contents.asArray() : [contents]) withoutQuotes(doc, doc.context.lookup(stream))
  clean(page.node.Resources())
}

/** Annotation types that are part of how the document works, not comments on it. */
const functionalAnnotations = new Set(['Link', 'Widget'])

/**
 * Acrobat's "Remove Hidden Information" for the places text can hide outside what the pages show:
 * document metadata, bookmarks, attached files, comments and mark-ups, hidden layers, and scripts
 * and automatic actions. Links and form fields keep working.
 */
function removeHiddenInformation(doc: PDFDocument): void {
  const info = doc.context.lookup(doc.context.trailerInfo.Info)
  if (info instanceof PDFDict) {
    for (const key of ['Title', 'Author', 'Subject', 'Keywords', 'Creator']) info.delete(PDFName.of(key))
  }
  // Attachments are listed in the /EmbeddedFiles name tree and, for PDF/A-3 and PDF 2.0, in /AF arrays.
  for (const key of ['Metadata', 'Outlines', 'OpenAction', 'AA', 'AF']) doc.catalog.delete(PDFName.of(key))
  const names = doc.catalog.lookupMaybe(PDFName.of('Names'), PDFDict)
  names?.delete(PDFName.of('EmbeddedFiles'))
  names?.delete(PDFName.of('JavaScript'))

  for (const page of doc.getPages()) {
    page.node.delete(PDFName.of('AA'))
    page.node.delete(PDFName.of('AF'))
    for (const ref of annotationRefs(page)) {
      const annotation = doc.context.lookup(ref)
      if (!(annotation instanceof PDFDict)) continue
      const subtype = annotation.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText() ?? ''
      const action = annotation.lookupMaybe(PDFName.of('A'), PDFDict)?.lookupMaybe(PDFName.of('S'), PDFName)?.decodeText()
      if (!functionalAnnotations.has(subtype) || action === 'JavaScript' || action === 'Launch') {
        page.node.removeAnnot(ref)
        continue
      }
      annotation.delete(PDFName.of('AA'))
    }
  }
  for (const field of doc.catalog.AcroForm() === undefined ? [] : doc.getForm().getFields()) {
    field.acroField.dict.delete(PDFName.of('AA'))
  }
  // Last, so bookmarks and actions that switch layers are already gone: this stops at anything still pointing to a layer.
  removeHiddenLayers(doc)
}

/**
 * Applies redaction boxes, given in displayed page coordinates (points from the top-left of the
 * page as shown, rotation included). Everything else on the page, and every other page, is kept.
 */
export async function redactPages(
  bytes: Uint8Array,
  requests: readonly RedactionRequest[],
  options: RedactOptions,
): Promise<Uint8Array> {
  const marked = await markRedactions(bytes, requests)
  const applied = await applyRedactions(marked, requests.map(request => request.index))

  const doc = await loadDocument(applied)
  const pages = doc.getPages()
  const redacted = requests.flatMap(request => pages[request.index] ?? [])
  redacted.forEach(page => removeMarkedText(doc, page))
  removeTagText(doc, new Set(redacted.map(page => page.ref.tag)))
  if (options.removeHiddenInformation) removeHiddenInformation(doc)
  dropUnreachable(doc)
  return doc.save({ updateFieldAppearances: false })
}
