import { PDFArray, PDFDict, PDFRef, PDFStream, type PDFDocument, type PDFObject } from '@cantoo/pdf-lib'

/**
 * Deletes every indirect object nothing in the document points to any more. pdf-lib writes
 * every object it holds, so without this a deleted page's content stream, or the appearance
 * stream of a redacted form field, would still sit in the saved file for anyone to extract.
 */
export function dropUnreachable(doc: PDFDocument): void {
  const { context } = doc
  const reached = new Set<string>()
  const pending: PDFObject[] = Object.values(context.trailerInfo).filter(entry => entry !== undefined)
  while (pending.length > 0) {
    const object = pending.pop()
    if (object instanceof PDFRef) {
      if (reached.has(object.tag)) continue
      reached.add(object.tag)
      const target = context.lookup(object)
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
