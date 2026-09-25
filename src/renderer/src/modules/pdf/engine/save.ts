import { degrees, PDFDocument, PDFName, PDFPageLeaf, type PDFPage } from '@cantoo/pdf-lib'
import { applyEdits, textFontProvider } from './draw'
import { applyFormValues } from './forms'
import { pageGeometry } from './geometry'
import { loadDocument } from './inspect'
import { dropUnreachable } from './prune'
import type { PageRef, SavePdfInput } from './types'

/**
 * Picks the `PDFPage` object for every output slot: source 0's own pages, and
 * copies of the other sources' pages. Each source page may appear once, because
 * a page object can only sit in the page tree once.
 */
async function resolvePages(
  doc: PDFDocument,
  sources: readonly Uint8Array[],
  pages: readonly PageRef[],
): Promise<PDFPage[]> {
  const originals = doc.getPages()
  const resolved = new Array<PDFPage>(pages.length)
  const listed = new Set<string>()
  /** Output positions of each other source's pages, copied in one `copyPages` call per source. */
  const copies = new Map<number, number[]>()

  pages.forEach((ref, position) => {
    if (sources[ref.source] === undefined) throw new Error(`No source document at index ${ref.source}`)
    const key = `${ref.source}:${ref.index}`
    if (listed.has(key)) throw new Error(`Page ${ref.index} of source ${ref.source} is listed twice`)
    listed.add(key)
    if (ref.source !== 0) {
      copies.set(ref.source, [...(copies.get(ref.source) ?? []), position])
      return
    }
    const page = originals[ref.index]
    if (page === undefined) throw new Error(`Source 0 has no page at index ${ref.index}`)
    resolved[position] = page
  })

  for (const [source, positions] of copies) {
    const donor = await loadDocument(sources[source])
    const donorPageCount = donor.getPageCount()
    const indexes = positions.map(position => {
      const index = pages[position].index
      if (index < 0 || index >= donorPageCount) throw new Error(`Source ${source} has no page at index ${index}`)
      return index
    })
    const copied = await doc.copyPages(donor, indexes)
    positions.forEach((position, slot) => {
      resolved[position] = copied[slot]
    })
  }

  return resolved
}

/**
 * Rebuilds the page list of source 0 in place, so its AcroForm, metadata,
 * attachments and outlines survive, then applies rotations, annotations and
 * form values in a single pass.
 *
 * Form fields that come in with pages copied from the other sources are not
 * registered in the output's AcroForm — they stay as plain widget annotations
 * and can only be flattened, not filled.
 */
export async function savePdf(input: SavePdfInput): Promise<Uint8Array> {
  const { sources, pages, formValues, flattenForm } = input
  if (sources.length === 0) throw new Error('savePdf needs at least one source document')
  if (pages.length === 0) throw new Error('savePdf needs at least one page')

  const doc = await loadDocument(sources[0])
  const originals = doc.getPages()
  const resolved = await resolvePages(doc, sources, pages)
  const kept = new Set(resolved)

  // A page can inherit its rotation, page boxes and resources from its parents in the page tree.
  // It is about to leave that tree, so it takes its own copy of each first; otherwise the saved
  // page would differ from the page shown, and redaction boxes would miss what they covered.
  // (`copyPages` does this for the pages from other sources.)
  for (const page of resolved) {
    for (const name of PDFPageLeaf.InheritableEntries) {
      const key = PDFName.of(name)
      const value = page.node.getInheritableAttribute(key)
      if (value !== undefined && page.node.get(key) === undefined) page.node.set(key, value)
    }
  }

  for (let index = originals.length - 1; index >= 0; index -= 1) doc.removePage(index)
  // `removePage` unregisters the page object; put back the ones we are reusing
  // so that every reference to them (annotations, outlines) stays valid.
  for (const page of originals) {
    if (kept.has(page)) doc.context.assign(page.ref, page.node)
  }
  resolved.forEach((page, index) => doc.insertPage(index, page))

  const font = textFontProvider(doc)
  for (let index = 0; index < pages.length; index += 1) {
    const ref = pages[index]
    const page = resolved[index]
    const geometry = pageGeometry(page, ref.rotation)
    if (geometry.rotation !== page.getRotation().angle) page.setRotation(degrees(geometry.rotation))
    if (ref.edits.length > 0) await applyEdits(doc, page, geometry, ref.edits, font)
  }

  const hasAcroForm = doc.catalog.AcroForm() !== undefined
  if (hasAcroForm || flattenForm === true) {
    const form = doc.getForm()
    if (formValues !== undefined) applyFormValues(doc, formValues)
    form.updateFieldAppearances()
    if (flattenForm === true) form.flatten({ updateFieldAppearances: false })
  }

  dropUnreachable(doc)
  return doc.save({ updateFieldAppearances: false })
}

/** Builds a new PDF holding only the given pages, in the given order. */
export async function extractPages(
  bytes: Uint8Array,
  indexes: readonly number[],
): Promise<Uint8Array> {
  if (indexes.length === 0) throw new Error('extractPages needs at least one page index')
  const source = await loadDocument(bytes)
  const pageCount = source.getPageCount()
  for (const index of indexes) {
    if (index < 0 || index >= pageCount) throw new Error(`No page at index ${index}`)
  }

  const target = await PDFDocument.create()
  const copied = await target.copyPages(source, [...indexes])
  for (const page of copied) target.addPage(page)
  return target.save()
}
