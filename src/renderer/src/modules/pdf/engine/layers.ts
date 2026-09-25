import { decodePDFRawStream, PDFArray, PDFDict, PDFName, PDFRawStream, PDFRef, type PDFDocument, type PDFObject } from '@cantoo/pdf-lib'
import { zlibSync } from 'fflate'
import { scanContent } from './content-stream'

/*
 * Optional content, which viewers call layers: content shown or hidden as a group. Acrobat's
 * "Remove hidden information" deletes what the default view hides and flattens the rest, so what
 * remains always shows. Hidden content sits in three places, and each is removed:
 *   - marked-content blocks `/OC /name BDC … EMC` in page and form XObject content,
 *   - XObjects (forms and images) whose own /OC is hidden, at every `Do` that paints them,
 *   - annotations whose /OC is hidden.
 * Then /OCProperties and the /OC entries go, and the blocks that stay become plain marked content.
 * Streams this does not read (annotation appearances, patterns, Type 3 glyphs) keep their content.
 */

const OC = PDFName.of('OC')
const XObject = PDFName.of('XObject')
const Properties = PDFName.of('Properties')

/** The group references in a value that is one group or an array of them; nulls are skipped. */
function groupRefs(doc: PDFDocument, value: PDFObject | undefined): PDFRef[] {
  const resolved = doc.context.lookup(value)
  if (resolved instanceof PDFArray) return resolved.asArray().filter((item): item is PDFRef => item instanceof PDFRef)
  return value instanceof PDFRef ? [value] : []
}

function typeOf(doc: PDFDocument, value: PDFObject | undefined): string | undefined {
  const dict = doc.context.lookup(value)
  return dict instanceof PDFDict ? dict.lookupMaybe(PDFName.of('Type'), PDFName)?.decodeText() : undefined
}

/**
 * Whether optional content (a group, or a membership dictionary over groups) is hidden in the
 * default configuration, /OCProperties /D: every group starts at /BaseState (ON unless given),
 * then the groups in /ON are on and those in /OFF are off.
 */
function hiddenTest(doc: PDFDocument, properties: PDFDict): (content: PDFObject | undefined) => boolean {
  const config = properties.lookupMaybe(PDFName.of('D'), PDFDict)
  const base = config?.lookupMaybe(PDFName.of('BaseState'), PDFName)?.decodeText() !== 'OFF'
  const on = new Map<string, boolean>()
  groupRefs(doc, properties.get(PDFName.of('OCGs'))).forEach(ref => on.set(ref.tag, base))
  groupRefs(doc, config?.get(PDFName.of('ON'))).forEach(ref => on.set(ref.tag, true))
  groupRefs(doc, config?.get(PDFName.of('OFF'))).forEach(ref => on.set(ref.tag, false))
  // A group the document does not list is under no configuration's control, so it shows.
  const groupOn = (value: PDFObject | undefined): boolean => !(value instanceof PDFRef) || (on.get(value.tag) ?? true)

  /** A visibility expression: a group, or `[/And …]`, `[/Or …]` or `[/Not …]` over expressions. */
  const expressionOn = (value: PDFObject | undefined, depth: number): boolean => {
    const expression = doc.context.lookup(value)
    if (!(expression instanceof PDFArray) || depth > 32) return groupOn(value)
    const [operator, ...operands] = expression.asArray()
    const results = operands.map(operand => expressionOn(operand, depth + 1))
    switch (operator instanceof PDFName ? operator.decodeText() : undefined) {
      case 'And': return results.every(Boolean)
      case 'Or': return results.some(Boolean)
      case 'Not': return !(results[0] ?? false)
      default: return true
    }
  }

  return content => {
    const dict = doc.context.lookup(content)
    if (!(dict instanceof PDFDict)) return false
    if (typeOf(doc, dict) !== 'OCMD') return !groupOn(content)
    // A membership dictionary: a visibility expression when it has one, else a policy over its groups.
    const expression = dict.get(PDFName.of('VE'))
    if (expression !== undefined) return !expressionOn(expression, 0)
    const groups = groupRefs(doc, dict.get(PDFName.of('OCGs'))).map(groupOn)
    if (groups.length === 0) return false
    switch (dict.lookupMaybe(PDFName.of('P'), PDFName)?.decodeText()) {
      case 'AllOn': return !groups.every(Boolean)
      case 'AnyOff': return groups.every(Boolean)
      case 'AllOff': return groups.some(Boolean)
      default: return !groups.some(Boolean)
    }
  }
}

function decodeContent(stream: PDFObject | undefined): Uint8Array {
  if (!(stream instanceof PDFRawStream)) throw new Error('A page has content that could not be read, so its hidden layers could not be removed.')
  try {
    return decodePDFRawStream(stream).decode()
  } catch {
    throw new Error('A page has content in an encoding that could not be read, so its hidden layers could not be removed.')
  }
}

/** A page's content, with /Contents given as an array joined into one stream: marked content can span them. */
function pageContent(doc: PDFDocument, contents: PDFObject | undefined): Uint8Array {
  const resolved = doc.context.lookup(contents)
  const streams = resolved instanceof PDFArray ? resolved.asArray() : resolved === undefined ? [] : [resolved]
  const parts = streams.map(stream => decodeContent(doc.context.lookup(stream)))
  return join(parts.flatMap((part, index) => index === 0 ? [part] : [newline, part]))
}

const newline = new Uint8Array([0x0a])

function join(parts: readonly Uint8Array[]): Uint8Array {
  const joined = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) {
    joined.set(part, offset)
    offset += part.length
  }
  return joined
}

/** Replaces a byte range; cut ranges become a space, so the tokens on either side stay apart. */
interface Edit {
  readonly start: number
  readonly end: number
  readonly replacement: string
}

/** Applies edits given in order and not overlapping. */
function applyEdits(content: Uint8Array, edits: readonly Edit[]): Uint8Array {
  const parts: Uint8Array[] = []
  let pos = 0
  for (const edit of edits) {
    parts.push(content.subarray(pos, edit.start), new TextEncoder().encode(edit.replacement))
    pos = edit.end
  }
  parts.push(content.subarray(pos))
  return join(parts)
}

/** XObject names each resource dictionary's content still paints, and names whose painting was removed. */
interface XObjectUse {
  readonly painted: Set<string>
  readonly dropped: Set<string>
}

/** Removes hidden optional content and flattens the rest; see the comment at the top of this file. */
export function removeHiddenLayers(doc: PDFDocument): void {
  const properties = doc.catalog.lookupMaybe(PDFName.of('OCProperties'), PDFDict)
  if (properties === undefined) return
  const isHidden = hiddenTest(doc, properties)
  const uses = new Map<PDFDict | undefined, XObjectUse>()
  const visitedForms = new Set<string>()

  /**
   * The edits for one content stream: a hidden `/OC … BDC` block goes with everything up to its
   * matching EMC (or the end of the stream), a `Do` of a hidden XObject goes, and a visible
   * `/OC … BDC` becomes a plain `/OC BMC` so nothing points to its layer any more.
   */
  const rewrite = (content: Uint8Array, resources: PDFDict | undefined): Uint8Array | undefined => {
    const named = (category: PDFName, name: string | undefined): PDFObject | undefined =>
      name === undefined ? undefined : resources?.lookupMaybe(category, PDFDict)?.get(PDFName.of(name.slice(1)))
    const use = uses.get(resources) ?? { painted: new Set<string>(), dropped: new Set<string>() }
    uses.set(resources, use)

    const edits: Edit[] = []
    let hiddenFrom: number | undefined
    let depth = 0
    for (const { operator, operands, start, end } of scanContent(content)) {
      const [first, second] = operands
      if (hiddenFrom !== undefined) {
        if (operator === 'BDC' || operator === 'BMC') depth += 1
        if (operator === 'EMC') depth -= 1
        if (operator === 'Do' && first !== undefined) use.dropped.add(first)
        if (depth === 0) {
          edits.push({ start: hiddenFrom, end, replacement: ' ' })
          hiddenFrom = undefined
        }
      } else if (operator === 'BDC' && first === '/OC') {
        if (isHidden(named(Properties, second))) {
          hiddenFrom = start
          depth = 1
        } else {
          edits.push({ start, end, replacement: '/OC BMC' })
        }
      } else if (operator === 'Do' && first !== undefined) {
        const xobject = doc.context.lookup(named(XObject, first))
        if (xobject instanceof PDFRawStream && isHidden(xobject.dict.get(OC))) {
          edits.push({ start, end, replacement: ' ' })
          use.dropped.add(first)
        } else {
          use.painted.add(first)
        }
      }
    }
    if (hiddenFrom !== undefined) edits.push({ start: hiddenFrom, end: content.length, replacement: ' ' })
    return edits.length === 0 ? undefined : applyEdits(content, edits)
  }

  /** Rewrites the form XObjects a resource dictionary names, and the forms inside those, once each. */
  const rewriteForms = (resources: PDFDict | undefined): void => {
    for (const [, value] of resources?.lookupMaybe(XObject, PDFDict)?.entries() ?? []) {
      if (!(value instanceof PDFRef) || visitedForms.has(value.tag)) continue
      visitedForms.add(value.tag)
      const form = doc.context.lookup(value)
      if (!(form instanceof PDFRawStream) || form.dict.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText() !== 'Form') continue
      if (isHidden(form.dict.get(OC))) continue
      // A form without its own resources uses those of whatever paints it.
      const own = form.dict.lookupMaybe(PDFName.of('Resources'), PDFDict) ?? resources
      const rewritten = rewrite(decodeContent(form), own)
      if (rewritten !== undefined) {
        form.dict.delete(PDFName.of('DecodeParms'))
        form.dict.set(PDFName.of('Filter'), PDFName.of('FlateDecode'))
        form.updateContents(zlibSync(rewritten))
      }
      rewriteForms(own)
    }
  }

  for (const page of doc.getPages()) {
    const resources = page.node.Resources()
    const rewritten = rewrite(pageContent(doc, page.node.get(PDFName.of('Contents'))), resources)
    if (rewritten !== undefined) page.node.set(PDFName.of('Contents'), doc.context.register(doc.context.flateStream(rewritten)))
    rewriteForms(resources)
    for (const ref of page.node.Annots()?.asArray() ?? []) {
      const annotation = doc.context.lookup(ref)
      if (!(ref instanceof PDFRef) || !(annotation instanceof PDFDict)) continue
      if (isHidden(annotation.get(OC))) page.node.removeAnnot(ref)
      else annotation.delete(OC)
    }
  }

  // XObjects painted only by removed content go with it, so their data leaves the file.
  for (const [resources, { painted, dropped }] of uses) {
    const xobjects = resources?.lookupMaybe(XObject, PDFDict)
    for (const [key, value] of xobjects?.entries() ?? []) {
      const xobject = doc.context.lookup(value)
      const name = `/${key.decodeText()}`
      if (!(xobject instanceof PDFRawStream)) continue
      if (isHidden(xobject.dict.get(OC)) || (dropped.has(name) && !painted.has(name))) xobjects?.delete(key)
      else xobject.dict.delete(OC)
    }
    const propertyLists = resources?.lookupMaybe(Properties, PDFDict)
    for (const [key, value] of propertyLists?.entries() ?? []) {
      const type = typeOf(doc, value)
      if (type === 'OCG' || type === 'OCMD') propertyLists?.delete(key)
    }
  }
  doc.catalog.delete(PDFName.of('OCProperties'))
}
