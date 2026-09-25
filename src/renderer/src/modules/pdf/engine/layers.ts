import { decodePDFRawStream, PDFArray, PDFContentStream, PDFDict, PDFName, PDFRawStream, PDFRef, type PDFDocument, type PDFObject } from '@cantoo/pdf-lib'
import { zlibSync } from 'fflate'
import { scanContent, type Instruction } from './content-stream'
import { removeWidgets } from './forms'
import { dropUnreachable } from './prune'

/*
 * Optional content, which viewers call layers: content shown or hidden as a group. Acrobat's
 * "Remove hidden information" deletes what the default view hides and flattens the rest, so what
 * remains always shows. Hidden content sits in these places, and each is removed:
 *   - marked-content blocks `/OC /name BDC … EMC` in every content stream the pages use: their
 *     own, and those of form XObjects, tiling patterns, Type 3 glyphs, soft masks and annotations,
 *   - XObjects (forms and images) whose own /OC is hidden, at every `Do` that paints them,
 *   - annotations whose /OC is hidden, and form fields whose widgets all are.
 * Then /OCProperties and the /OC entries go, and the blocks that stay become plain marked content.
 * Without /OCProperties every layer shows. So if anything still points to a layer at the end,
 * hidden content sits somewhere this did not clean, and removal stops with an error.
 */

const OC = PDFName.of('OC')
const XObject = PDFName.of('XObject')
const Properties = PDFName.of('Properties')
const Resources = PDFName.of('Resources')

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
  const unreadable = (): Error => unsafe('A page has content Zendo can’t read, so it can’t remove the hidden layers.')
  // pdf-lib wraps a page's content in `q` … `Q` streams of its own the first time it changes the page.
  if (stream instanceof PDFContentStream) return stream.getUnencodedContents()
  if (!(stream instanceof PDFRawStream)) throw unreadable()
  try {
    return decodePDFRawStream(stream).decode()
  } catch {
    throw unreadable()
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

/**
 * The resources each resource dictionary's content still uses, and those that removed content
 * used, each as a category and name such as `Pattern /P1`. Something only removed content used
 * goes with it, so its data leaves the file.
 */
interface ResourceUse {
  readonly used: Set<string>
  readonly dropped: Set<string>
}

/** The resource categories whose entries go when only removed content used them. */
const droppable = ['XObject', 'Pattern', 'Shading', 'ExtGState']

/** The resource an instruction names: `Do` an XObject, `sh` a shading, `scn` a pattern to paint with, `gs` a graphics state. */
function resourceOf({ operator, operands }: Instruction): string | undefined {
  const name = operands.at(-1)
  if (name === undefined || !name.startsWith('/')) return undefined
  if (operator === 'Do') return `XObject ${name}`
  if (operator === 'sh') return `Shading ${name}`
  if (operator === 'scn' || operator === 'SCN') return `Pattern ${name}`
  if (operator === 'gs') return `ExtGState ${name}`
  return undefined
}

/**
 * Of the graphics state, what refers to resources: the instructions that set the current fill and
 * stroke patterns and the graphics states in effect. Whatever paints next paints with these.
 */
interface ResourceState {
  readonly fill?: number
  readonly stroke?: number
  readonly states: readonly number[]
}

/** An error that says what to do instead, for content this cannot clean without changing what shows. */
function unsafe(reason: string): Error {
  return new Error(`${reason} Save without “Also remove hidden information”, or remove the layer in Adobe Acrobat.`)
}

function scan(content: Uint8Array): Instruction[] {
  try {
    return scanContent(content)
  } catch {
    throw unsafe('This PDF has an image in its page content that Zendo can’t read past, so it can’t remove the hidden layers safely.')
  }
}

/*
 * Hidden content still changes state that visible content after it relies on: viewers skip only
 * its painting. Its `cm`, colours, line and text settings and clip (`W n`) stay in effect, and its
 * text shows move the text position. A hidden block goes whole only when none of that reaches past
 * it; otherwise it keeps every state change and loses only what paints.
 */

/** Operators whose effect lasts until a `Q` restores the graphics state: the CTM, colours, line and text settings, and the clip. */
const stateOperators = new Set(['cm', 'w', 'J', 'j', 'M', 'd', 'ri', 'i', 'gs', 'CS', 'cs', 'SC', 'SCN', 'sc', 'scn', 'G', 'g', 'RG', 'rg', 'K', 'k', 'Tc', 'Tw', 'Tz', 'TL', 'Tf', 'Tr', 'Ts', '"', 'W', 'W*'])
/** Operators that set or advance the text position, which lasts until the text object's `ET`. */
const textOperators = new Set(['Td', 'TD', 'Tm', 'T*', 'Tj', 'TJ', "'", '"'])
const textShows = new Set(['Tj', 'TJ', "'", '"'])
/** Operators after which text is placed from the start of a line, so a show removed before them moves nothing. */
const lineStarts = new Set(['BT', 'ET', 'Td', 'TD', 'Tm', 'T*', "'", '"'])
const pathPainting = new Set(['S', 's', 'f', 'F', 'f*', 'B', 'B*', 'b', 'b*'])
/** What paints, and so uses the patterns and graphics states in effect. A form painted by `Do` inherits them. */
const painting = new Set([...pathPainting, ...textShows, 'Do', 'BI', 'sh'])
const fillColours = new Set(['cs', 'sc', 'scn', 'g', 'rg', 'k'])
const strokeColours = new Set(['CS', 'SC', 'SCN', 'G', 'RG', 'K'])
/** What a hidden block that cannot go whole keeps as written: everything that changes state and paints nothing. */
const kept = new Set([
  ...stateOperators, 'q', 'Q', 'BT', 'ET', 'Td', 'TD', 'Tm', 'T*',
  'm', 'l', 'c', 'v', 'y', 'h', 're', 'n', 'BMC', 'EMC', 'MP', 'BX', 'EX', 'd0', 'd1',
])

/** The index of the `EMC` that closes the marked content opened at `open`, or the end when none does. */
function closingIndex(instructions: readonly Instruction[], open: number): number {
  let depth = 0
  for (let index = open; index < instructions.length; index += 1) {
    const { operator } = instructions[index]
    if (operator === 'BDC' || operator === 'BMC') depth += 1
    if (operator === 'EMC') depth -= 1
    if (depth === 0) return index
  }
  return instructions.length
}

/**
 * Whether a hidden block can go whole: `q`/`Q` and `BT`/`ET` balance inside it, every state change
 * sits inside a `q … Q` of its own, and it moves the text position only inside a text object of its
 * own. Text settings such as the font are graphics state, so a `Tf` outlasts `ET`; only the text
 * position ends there.
 */
function selfContained(block: readonly Instruction[], insideText: boolean): boolean {
  let saves = 0
  let text = false
  for (const { operator } of block) {
    if (operator === 'q') saves += 1
    if (operator === 'Q' && --saves < 0) return false
    // A text object inside another is not valid PDF, and its `BT` would reset the outer one's position.
    if (operator === 'BT' && (text || insideText)) return false
    if (operator === 'BT') text = true
    if (operator === 'ET' && !text) return false
    if (operator === 'ET') text = false
    if (textOperators.has(operator) && !text) return false
    if (saves === 0 && stateOperators.has(operator)) return false
  }
  return saves === 0 && !text
}

/**
 * An instruction of a hidden block that cannot go whole, without its painting: `undefined` keeps it
 * as written, a string replaces it. Paths end unpainted (`n` still applies a pending clip), `'` and
 * `"` still move to the next line, and marked content loses its property list, which can quote the
 * hidden text as /ActualText. Text shows, images, shadings, XObjects and unknown operators go.
 */
function withoutPainting({ operator, operands: [first, second] }: Instruction): string | undefined {
  if (operator === "'") return 'T*'
  if (operator === '"') return `${first ?? '0'} Tw ${second ?? '0'} Tc T*`
  if (kept.has(operator)) return undefined
  if (pathPainting.has(operator)) return 'n'
  const tag = first !== undefined && /^\/[\w.-]+$/.test(first) ? first : '/Span'
  if (operator === 'BDC') return `${tag} BMC`
  if (operator === 'DP') return `${tag} MP`
  return ' '
}

/** Whether every opening operator has its closing one after it, as in `q … Q`. */
function balanced(instructions: readonly Instruction[], opens: readonly string[], close: string): boolean {
  let depth = 0
  for (const { operator } of instructions) {
    if (opens.includes(operator)) depth += 1
    if (operator === close && --depth < 0) return false
  }
  return depth === 0
}

/** Removes hidden optional content and flattens the rest; see the comment at the top of this file. */
export function removeHiddenLayers(doc: PDFDocument): void {
  const properties = doc.catalog.lookupMaybe(PDFName.of('OCProperties'), PDFDict)
  if (properties === undefined) return
  const isHidden = hiddenTest(doc, properties)
  const uses = new Map<PDFDict | undefined, ResourceUse>()
  const visited = new Set<string>()

  /**
   * The edits for one content stream: a hidden `/OC … BDC` block goes with everything up to its
   * matching EMC (or the end of the stream), or when that would change what follows it, loses only
   * its painting (see `selfContained` and `withoutPainting`). A `Do` of a hidden XObject goes, and
   * an `/OC … BDC` that stays becomes a plain `/OC BMC`, so nothing points to its layer any more.
   */
  const rewrite = (content: Uint8Array, resources: PDFDict | undefined): Uint8Array | undefined => {
    const named = (category: PDFName, name: string | undefined): PDFObject | undefined =>
      name === undefined ? undefined : resources?.lookupMaybe(category, PDFDict)?.get(PDFName.of(name.slice(1)))
    const use = uses.get(resources) ?? { used: new Set<string>(), dropped: new Set<string>() }
    uses.set(resources, use)

    const instructions = scan(content)
    const edits: Edit[] = []
    /** Instructions before this index are in a hidden block that loses its painting. */
    let hiddenEnd = -1
    let inText = false
    /** A hidden text show was removed and nothing has placed text since: visible text shown now would move. */
    let shifted = false
    let state: ResourceState = { states: [] }
    const saved: ResourceState[] = []
    /** Instructions that set a pattern or graphics state something visible then painted with. */
    const paintedWith = new Set<number>()
    /** Such instructions in emptied hidden blocks: each stays only if something visible painted with it. */
    const setters: number[] = []
    for (let index = 0; index < instructions.length; index += 1) {
      const instruction = instructions[index]
      const { operator, operands: [first, second], start, end } = instruction
      const resource = resourceOf(instruction)
      const hidden = index < hiddenEnd
      if (operator === 'BT') inText = true
      if (operator === 'ET') inText = false
      if (lineStarts.has(operator)) shifted = false
      if (operator === 'q') saved.push(state)
      if (operator === 'Q') state = saved.pop() ?? state
      if (fillColours.has(operator)) state = { ...state, fill: resource === undefined ? undefined : index }
      if (strokeColours.has(operator)) state = { ...state, stroke: resource === undefined ? undefined : index }
      if (operator === 'gs') state = { ...state, states: [...state.states, index] }
      if (!hidden && painting.has(operator)) {
        for (const setter of [state.fill, state.stroke, ...state.states]) if (setter !== undefined) paintedWith.add(setter)
      }

      if (hidden) {
        const replacement = withoutPainting(instruction)
        if (replacement !== undefined) edits.push({ start, end, replacement })
        if (textShows.has(operator)) shifted = true
        if (resource !== undefined && replacement !== undefined) use.dropped.add(resource)
        if (resource !== undefined && replacement === undefined) setters.push(index)
      } else if (shifted && (operator === 'Tj' || operator === 'TJ')) {
        throw unsafe('This PDF mixes a hidden layer into visible text in a way Zendo can’t remove safely.')
      } else if (operator === 'BDC' && first === '/OC') {
        if (isHidden(named(Properties, second))) {
          const close = closingIndex(instructions, index)
          const block = instructions.slice(index + 1, close)
          if (selfContained(block, inText)) {
            edits.push({ start, end: instructions[close]?.end ?? content.length, replacement: ' ' })
            for (const inner of block) {
              const used = resourceOf(inner)
              if (used !== undefined) use.dropped.add(used)
            }
            index = close
            continue
          }
          hiddenEnd = close
        }
        edits.push({ start, end, replacement: '/OC BMC' })
      } else if (resource !== undefined) {
        const xobject = operator === 'Do' ? doc.context.lookup(named(XObject, first)) : undefined
        if (xobject instanceof PDFRawStream && isHidden(xobject.dict.get(OC))) {
          edits.push({ start, end, replacement: ' ' })
          use.dropped.add(resource)
        } else {
          use.used.add(resource)
        }
      }
    }
    // A pattern or graphics state set in an emptied block, that nothing visible paints with before
    // it is replaced or restored, changes nothing: it goes, and its resource can go with it.
    for (const index of setters) {
      const setter = instructions[index]
      const resource = resourceOf(setter)
      if (resource === undefined) continue
      if (paintedWith.has(index)) {
        use.used.add(resource)
      } else {
        edits.push({ start: setter.start, end: setter.end, replacement: ' ' })
        use.dropped.add(resource)
      }
    }
    if (edits.length === 0) return undefined
    edits.sort((a, b) => a.start - b.start)
    const output = applyEdits(content, edits)
    // A check on the above: emptying a block keeps its `q`/`Q` and marked-content pairs, so their balance must not change.
    const after = scan(output)
    for (const [opens, close] of [[['q'], 'Q'], [['BDC', 'BMC'], 'EMC']] as const) {
      if (balanced(instructions, opens, close) && !balanced(after, opens, close)) {
        throw unsafe('Zendo couldn’t remove this PDF’s hidden layers without changing what its pages show.')
      }
    }
    return output
  }

  /** Rewrites a content stream other than a page's own, then the streams its resources name; each stream once. */
  const rewriteStream = (value: PDFObject | undefined, inherited: PDFDict | undefined): void => {
    if (!(value instanceof PDFRef) || visited.has(value.tag)) return
    visited.add(value.tag)
    const stream = doc.context.lookup(value)
    if (!(stream instanceof PDFRawStream) || isHidden(stream.dict.get(OC))) return
    // A stream without resources of its own uses those of whatever paints it.
    const resources = stream.dict.lookupMaybe(Resources, PDFDict) ?? inherited
    const rewritten = rewrite(decodeContent(stream), resources)
    if (rewritten !== undefined) {
      stream.dict.delete(PDFName.of('DecodeParms'))
      stream.dict.set(PDFName.of('Filter'), PDFName.of('FlateDecode'))
      stream.updateContents(zlibSync(rewritten))
    }
    rewriteNamed(resources)
  }

  /** Rewrites the content streams a resource dictionary names: forms, tiling patterns, Type 3 glyphs and soft masks. */
  const rewriteNamed = (resources: PDFDict | undefined): void => {
    const named = (category: string): PDFObject[] => resources?.lookupMaybe(PDFName.of(category), PDFDict)?.values() ?? []
    for (const value of named('XObject')) {
      const xobject = doc.context.lookup(value)
      if (xobject instanceof PDFRawStream && xobject.dict.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText() === 'Form') rewriteStream(value, resources)
    }
    // Tiling patterns are streams; shading patterns are dictionaries without content.
    for (const value of named('Pattern')) if (doc.context.lookup(value) instanceof PDFRawStream) rewriteStream(value, resources)
    for (const value of named('Font')) {
      const font = doc.context.lookup(value)
      if (!(font instanceof PDFDict)) continue
      const own = font.lookupMaybe(Resources, PDFDict) ?? resources
      for (const glyph of font.lookupMaybe(PDFName.of('CharProcs'), PDFDict)?.values() ?? []) rewriteStream(glyph, own)
    }
    for (const value of named('ExtGState')) {
      const state = doc.context.lookup(value)
      const mask = state instanceof PDFDict ? doc.context.lookup(state.get(PDFName.of('SMask'))) : undefined
      if (mask instanceof PDFDict) rewriteStream(mask.get(PDFName.of('G')), resources)
    }
  }

  // Before the pages drop their hidden annotations, so each field still finds its widgets' pages.
  removeWidgets(doc, widget => isHidden(widget.dict.get(OC)))

  for (const page of doc.getPages()) {
    const resources = page.node.Resources()
    const rewritten = rewrite(pageContent(doc, page.node.get(PDFName.of('Contents'))), resources)
    if (rewritten !== undefined) page.node.set(PDFName.of('Contents'), doc.context.register(doc.context.flateStream(rewritten)))
    rewriteNamed(resources)
    for (const ref of page.node.Annots()?.asArray() ?? []) {
      const annotation = doc.context.lookup(ref)
      if (!(ref instanceof PDFRef) || !(annotation instanceof PDFDict)) continue
      if (isHidden(annotation.get(OC))) {
        page.node.removeAnnot(ref)
        continue
      }
      annotation.delete(OC)
      // An action that turns layers on or off has none left to turn.
      if (annotation.lookupMaybe(PDFName.of('A'), PDFDict)?.lookupMaybe(PDFName.of('S'), PDFName)?.decodeText() === 'SetOCGState') annotation.delete(PDFName.of('A'))
      // Each appearance is a form, or a dictionary of forms by state, such as a checkbox's on and off.
      for (const value of annotation.lookupMaybe(PDFName.of('AP'), PDFDict)?.values() ?? []) {
        const states = doc.context.lookup(value)
        for (const form of states instanceof PDFDict ? states.values() : [value]) rewriteStream(form, undefined)
      }
    }
  }

  for (const [resources, { used, dropped }] of uses) {
    // Resources only removed content used go with it, so their data leaves the file.
    for (const category of droppable) {
      const entries = resources?.lookupMaybe(PDFName.of(category), PDFDict)
      for (const [key] of entries?.entries() ?? []) {
        const resource = `${category} /${key.decodeText()}`
        if (dropped.has(resource) && !used.has(resource)) entries?.delete(key)
      }
    }
    const xobjects = resources?.lookupMaybe(XObject, PDFDict)
    for (const [key, value] of xobjects?.entries() ?? []) {
      const xobject = doc.context.lookup(value)
      if (!(xobject instanceof PDFRawStream)) continue
      if (isHidden(xobject.dict.get(OC))) xobjects?.delete(key)
      else xobject.dict.delete(OC)
    }
    const propertyLists = resources?.lookupMaybe(Properties, PDFDict)
    for (const [key, value] of propertyLists?.entries() ?? []) {
      const type = typeOf(doc, value)
      if (type === 'OCG' || type === 'OCMD') propertyLists?.delete(key)
    }
  }
  doc.catalog.delete(PDFName.of('OCProperties'))

  dropUnreachable(doc)
  if (doc.context.enumerateIndirectObjects().some(([, object]) => typeOf(doc, object) === 'OCG')) {
    throw unsafe('This PDF keeps part of a hidden layer in a place Zendo can’t clean.')
  }
}
