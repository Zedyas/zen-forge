import { PDFArray, PDFDict, PDFName, PDFRef, PDFStream, type PDFDocument, type PDFObject } from '@cantoo/pdf-lib'

function isPageNode(object: PDFObject | undefined): boolean {
  const type = object instanceof PDFDict ? object.lookupMaybe(PDFName.of('Type'), PDFName)?.decodeText() : undefined
  return type === 'Page' || type === 'Pages'
}

/** The references of the page tree's nodes: the pages the document shows, and the nodes above them. */
function pageTree(doc: PDFDocument): Set<string> {
  const tree = new Set<string>()
  const pending: (PDFObject | undefined)[] = [doc.catalog.get(PDFName.of('Pages'))]
  while (pending.length > 0) {
    const ref = pending.pop()
    if (!(ref instanceof PDFRef) || tree.has(ref.tag)) continue
    tree.add(ref.tag)
    const kids = doc.context.lookup(ref, PDFDict).lookupMaybe(PDFName.of('Kids'), PDFArray)
    pending.push(...(kids?.asArray() ?? []))
  }
  return tree
}

/**
 * Deletes every indirect object nothing in the document points to any more. pdf-lib writes
 * every object it holds, so without this a deleted page's content stream, or the appearance
 * stream of a redacted form field, would still sit in the saved file for anyone to extract.
 * A page outside the page tree counts as unreachable even when a link, bookmark or form field
 * still points to it: a deleted page, or one `copyPages` brought along with such a pointer.
 */
export function dropUnreachable(doc: PDFDocument): void {
  const { context } = doc
  const tree = pageTree(doc)
  const reached = new Set<string>()
  const pending: PDFObject[] = Object.values(context.trailerInfo).filter(entry => entry !== undefined)
  while (pending.length > 0) {
    const object = pending.pop()
    if (object instanceof PDFRef) {
      if (reached.has(object.tag)) continue
      const target = context.lookup(object)
      if (isPageNode(target) && !tree.has(object.tag)) continue
      reached.add(object.tag)
      if (target !== undefined) pending.push(target)
    } else if (object instanceof PDFDict) {
      pending.push(...object.values())
    } else if (object instanceof PDFArray) {
      pending.push(...object.asArray())
    } else if (object instanceof PDFStream) {
      pending.push(object.dict)
    }
  }
  for (const [ref] of context.enumerateIndirectObjects()) {
    if (!reached.has(ref.tag)) context.delete(ref)
  }
}
