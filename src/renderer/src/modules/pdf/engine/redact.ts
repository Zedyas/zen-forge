import {
  PDFArray,
  PDFDict,
  PDFName,
  PDFNumber,
  PDFRef,
  type PDFDocument,
  type PDFObject,
  type PDFPage,
} from '@cantoo/pdf-lib'
import { pageGeometry, rectToUserSpace, type Rect } from './geometry'
import { loadDocument } from './inspect'
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

/** Removes form fields (value included) whose widgets all touch a box, and the touching widgets of the rest. */
function removeTouchedFields(doc: PDFDocument, areas: ReadonlyMap<PDFPage, readonly Rect[]>): void {
  if (doc.catalog.AcroForm() === undefined) return
  const form = doc.getForm()
  const pageOfWidget = new Map<string, PDFPage>()
  for (const page of areas.keys()) annotationRefs(page).forEach(ref => pageOfWidget.set(ref.tag, page))

  for (const field of form.getFields()) {
    const widgets = field.acroField.getWidgets()
    const doomed = widgets.map(widget => {
      const ref = doc.context.getObjectRef(widget.dict)
      const page = ref === undefined ? undefined : pageOfWidget.get(ref.tag)
      return page !== undefined && touchesAny(widget.getRectangle(), areas.get(page) ?? [])
    })
    if (!doomed.includes(true)) continue
    if (doomed.every(Boolean)) {
      form.removeField(field)
      continue
    }
    for (let index = widgets.length - 1; index >= 0; index -= 1) {
      const widget = widgets[index]
      const ref = widget === undefined ? undefined : doc.context.getObjectRef(widget.dict)
      if (!doomed[index] || ref === undefined) continue
      pageOfWidget.get(ref.tag)?.node.removeAnnot(ref)
      field.acroField.removeWidget(index)
    }
  }
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
  for (const { index, boxes } of requests) {
    const page = pages[index]
    if (page === undefined) throw new Error(`No page at index ${index}`)
    const geometry = pageGeometry(page)
    areas.set(page, boxes.map(box => rectToUserSpace(geometry, box)))
  }

  removeTouchedFields(doc, areas)
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
    const ownPage = object.get(PDFName.of('Pg'))
    const page = ownPage instanceof PDFRef ? ownPage.tag : inheritedPage
    if (page !== undefined && redactedPages.has(page)) {
      for (const key of ['Alt', 'ActualText', 'E', 'T']) object.delete(PDFName.of(key))
    }
    visit(object.get(PDFName.of('K')), page)
  }
  visit(root.get(PDFName.of('K')), undefined)
}

/** Annotation types that are part of how the document works, not comments on it. */
const functionalAnnotations = new Set(['Link', 'Widget'])

/**
 * Acrobat's "Remove Hidden Information" for the places text can hide outside the pages: document
 * metadata, bookmarks, attached files, comments and mark-ups, and scripts and automatic actions.
 * Links and form fields keep working.
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
  const redactedPages = new Set(requests.flatMap(request => pages[request.index]?.ref.tag ?? []))
  removeTagText(doc, redactedPages)
  if (options.removeHiddenInformation) removeHiddenInformation(doc)
  dropUnreachable(doc)
  return doc.save({ updateFieldAppearances: false })
}
