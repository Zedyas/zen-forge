import {
  decodePDFRawStream,
  PDFArray,
  PDFContentStream,
  PDFDict,
  PDFHexString,
  PDFName,
  PDFNull,
  PDFNumber,
  PDFOperator,
  PDFOperatorNames,
  PDFRawStream,
  PDFRef,
  PDFStream,
  PDFString,
  beginMarkedContent,
  type PDFContext,
  type PDFDocument,
  type PDFObject,
  type PDFObjectCopier,
  type PDFPage,
} from '@cantoo/pdf-lib'

/*
 * Tagged PDF in short: the catalog's /StructTreeRoot holds a tree of structure elements
 * (/S is the role, such as /P or /Figure) in reading order. An element's /K lists its kids:
 * other elements, and the content they stand for. Content is either a marked-content id
 * (MCID: a number set by `/P <</MCID 3>> BDC … EMC` in a page's content stream, given as a
 * bare number or an /MCR dictionary) or an /OBJR dictionary that points to an annotation such
 * as a form field. /Pg names the page the content is on. The /ParentTree goes the other way,
 * from content to element: each page's /StructParents key holds an array indexed by MCID, and
 * each annotation's /StructParent key holds its element. Screen readers read the tree in order.
 */

const K = PDFName.of('K')
const Pg = PDFName.of('Pg')
const P = PDFName.of('P')
const S = PDFName.of('S')
const Obj = PDFName.of('Obj')
const MCID = PDFName.of('MCID')
const Stm = PDFName.of('Stm')
const StructParents = PDFName.of('StructParents')
const StructParent = PDFName.of('StructParent')
const StructTreeRoot = PDFName.of('StructTreeRoot')

/**
 * The highest MCID put in the parent tree, whose array for a page has one slot per MCID up to
 * the highest. Real pages use thousands at most; a crafted /MCID 999999999 would otherwise
 * make an array of a billion slots.
 */
const highestMcid = 1_000_000

/** Standard roles that only group other elements, so new content can go inside them. */
const groupingRoles = new Set(['Document', 'DocumentFragment', 'Part', 'Art', 'Sect', 'Div', 'NonStruct'])

type Kid =
  | { readonly kind: 'mcid'; readonly page: string | undefined; readonly stream: PDFRef | undefined; readonly mcid: number }
  | { readonly kind: 'object'; readonly page: string | undefined; readonly target: PDFRef | undefined }
  | { readonly kind: 'element'; readonly dict: PDFDict }

function rootOf(doc: PDFDocument): PDFDict | undefined {
  return doc.catalog.lookupMaybe(StructTreeRoot, PDFDict)
}

function refTag(object: PDFObject | undefined): string | undefined {
  return object instanceof PDFRef ? object.tag : undefined
}

function kidsOf(node: PDFDict): PDFObject[] {
  const k = node.get(K)
  if (k === undefined) return []
  const resolved = node.context.lookup(k)
  return resolved instanceof PDFArray ? resolved.asArray() : [k]
}

function setKids(node: PDFDict, kids: readonly PDFObject[]): void {
  node.set(K, node.context.obj([...kids]))
}

/** What one entry of an element's /K is. `page` is the entry's own /Pg, or the element's. */
function readKid(context: PDFContext, kid: PDFObject, page: string | undefined): Kid | undefined {
  if (kid instanceof PDFNumber) return { kind: 'mcid', page, stream: undefined, mcid: kid.asNumber() }
  const dict = context.lookup(kid)
  if (!(dict instanceof PDFDict)) return undefined
  const ownPage = refTag(dict.get(Pg)) ?? page
  const mcid = dict.lookup(MCID)
  if (mcid instanceof PDFNumber) {
    const stream = dict.get(Stm)
    return { kind: 'mcid', page: ownPage, stream: stream instanceof PDFRef ? stream : undefined, mcid: mcid.asNumber() }
  }
  if (dict.has(Obj)) {
    const target = dict.get(Obj)
    return { kind: 'object', page: ownPage, target: target instanceof PDFRef ? target : undefined }
  }
  return { kind: 'element', dict }
}

function pageOf(element: PDFDict, inherited: string | undefined): string | undefined {
  return refTag(element.get(Pg)) ?? inherited
}

/** The standard role of an element, through the root's /RoleMap. */
function standardRole(root: PDFDict, element: PDFDict): string | undefined {
  const roleMap = root.lookupMaybe(PDFName.of('RoleMap'), PDFDict)
  let role = element.lookup(S)
  for (let hop = 0; hop < 10 && role instanceof PDFName && roleMap?.has(role) === true; hop += 1) role = roleMap.lookup(role)
  return role instanceof PDFName ? role.decodeText() : undefined
}

/** The pages an element's content is on, in its whole subtree. */
function contentPages(context: PDFContext, element: PDFDict, inherited: string | undefined, visited = new Set<PDFDict>()): Set<string> {
  const pages = new Set<string>()
  if (visited.has(element)) return pages
  visited.add(element)
  const page = pageOf(element, inherited)
  for (const kid of kidsOf(element)) {
    const read = readKid(context, kid, page)
    if (read === undefined) continue
    if (read.kind === 'element') contentPages(context, read.dict, page, visited).forEach(tag => pages.add(tag))
    else if (read.page !== undefined) pages.add(read.page)
  }
  return pages
}

export interface LiveContent {
  /** Tags of the pages in the document's page list. */
  readonly pages: ReadonlySet<string>
  /** Tags of the annotations on those pages. */
  readonly annotations: ReadonlySet<string>
}

export function liveContent(doc: PDFDocument, pages: readonly PDFPage[] = doc.getPages()): LiveContent {
  return {
    pages: new Set(pages.map(page => page.ref.tag)),
    annotations: new Set(pages.flatMap(page => page.node.Annots()?.asArray() ?? []).flatMap(ref => ref instanceof PDFRef ? [ref.tag] : [])),
  }
}

/**
 * Removes the content entries for pages and annotations that are gone, and the elements that
 * are left with no content. An element that had no content to begin with stays, unless it
 * names a page that is gone: its alt text can describe that page.
 * Returns whether `element` still has content.
 */
function prune(context: PDFContext, element: PDFDict, inherited: string | undefined, live: LiveContent, visited: Set<PDFDict>): boolean {
  if (visited.has(element)) return false
  visited.add(element)
  const page = pageOf(element, inherited)
  const kids = kidsOf(element)
  const kept = kids.filter(kid => {
    const read = readKid(context, kid, page)
    if (read === undefined) return false
    if (read.kind === 'element') return prune(context, read.dict, page, live, visited)
    if (read.kind === 'object') {
      const target = read.target === undefined ? undefined : context.lookup(read.target)
      // An annotation goes with its page, or alone (a deleted or flattened form field).
      const isAnnotation = target instanceof PDFDict && target.has(PDFName.of('Subtype')) && target.has(PDFName.of('Rect'))
      if (isAnnotation) return read.target !== undefined && live.annotations.has(read.target.tag)
    }
    return read.page !== undefined && live.pages.has(read.page)
  })
  if (kept.length !== kids.length) setKids(element, kept)
  return kept.length > 0 || (kids.length === 0 && (page === undefined || live.pages.has(page)))
}

/** Removes the tags of content that is no longer in the document. The root itself always stays. */
export function pruneStructure(doc: PDFDocument, live: LiveContent): void {
  const root = rootOf(doc)
  if (root !== undefined) prune(doc.context, root, undefined, live, new Set())
}

/**
 * Puts every element's kids in the order of the pages their content starts on. Kids on the
 * same page keep their order, and a kid with no content stays after the kid before it, so
 * within a page the original reading order stays. Run only when the page order changed:
 * a tree may read content in an order the pages don't have (a sidebar after the main text).
 */
export function orderStructureByPages(doc: PDFDocument): void {
  const root = rootOf(doc)
  if (root === undefined) return
  const { context } = doc
  const position = new Map(doc.getPages().map((page, index) => [page.ref.tag, index]))
  const visited = new Set<PDFDict>()
  const firstPosition = (pages: Iterable<string>): number => Math.min(...[...pages].map(tag => position.get(tag) ?? Infinity))

  const order = (element: PDFDict, inherited: string | undefined): void => {
    if (visited.has(element)) return
    visited.add(element)
    const page = pageOf(element, inherited)
    const kids = kidsOf(element)
    let previous = -Infinity
    const keyed = kids.map((kid, index) => {
      const read = readKid(context, kid, page)
      let key = Infinity
      if (read?.kind === 'element') {
        key = firstPosition(contentPages(context, read.dict, page))
        order(read.dict, page)
      } else if (read !== undefined && read.page !== undefined) {
        key = position.get(read.page) ?? Infinity
      }
      if (key === Infinity) key = previous
      previous = key
      return { kid, key, index }
    })
    keyed.sort((a, b) => a.key - b.key || a.index - b.index)
    if (keyed.some((entry, index) => entry.index !== index)) setKids(element, keyed.map(entry => entry.kid))
  }
  order(root, undefined)
}

function ensureRoot(doc: PDFDocument): PDFDict {
  const existing = rootOf(doc)
  if (existing !== undefined) return existing
  const root = doc.context.obj({ Type: 'StructTreeRoot', K: [] })
  doc.catalog.set(StructTreeRoot, doc.context.register(root))
  return root
}

/** The root's reference, registering it first if the file wrote it inline. */
function rootRef(doc: PDFDocument, root: PDFDict): PDFRef {
  const entry = doc.catalog.get(StructTreeRoot)
  if (entry instanceof PDFRef) return entry
  const ref = doc.context.register(root)
  doc.catalog.set(StructTreeRoot, ref)
  return ref
}

/** The tree's one /Document element, when the root holds only that. */
function soleDocument(context: PDFContext, root: PDFDict): { readonly element: PDFDict; readonly ref: PDFRef } | undefined {
  const kids = kidsOf(root)
  const only = kids.length === 1 ? kids[0] : undefined
  const element = only === undefined ? undefined : context.lookup(only)
  if (!(only instanceof PDFRef) || !(element instanceof PDFDict)) return undefined
  return standardRole(root, element) === 'Document' ? { element, ref: only } : undefined
}

/**
 * Copies `donor`'s tags for the pages `copied` (already copied into `doc` with the same
 * `copier`, so each /Pg maps to the copy) and appends them to `doc`'s tree. The donor is
 * changed: its tree is first cut down to those pages, so tags of pages left behind, which can
 * quote their text, never reach the output. Returns whether the donor had tags.
 */
export function copyStructure(doc: PDFDocument, donor: PDFDocument, donorPages: readonly PDFPage[], copier: PDFObjectCopier): boolean {
  const donorRoot = rootOf(donor)
  if (donorRoot === undefined) return false
  pruneStructure(donor, liveContent(donor, donorPages))
  const root = ensureRoot(doc)
  const ownDocument = soleDocument(doc.context, root)
  const donorDocument = soleDocument(donor.context, donorRoot)
  // Sections can only be put in page order among the kids of one element, so the donor's
  // sections go into this file's /Document, unwrapped from the donor's own. A file with no
  // tags of its own takes the donor's tree as it is.
  const fresh = kidsOf(root).length === 0
  const unwrap = !fresh && donorDocument !== undefined
  const kids = kidsOf(unwrap ? donorDocument.element : donorRoot)
  const into = ownDocument ?? { element: root, ref: rootRef(doc, root) }
  const donorElements = elementsUnder(donor.context, donorRoot)
  const roleRenames = renameClashes(doc.context, root, donor.context, donorRoot, 'RoleMap')
  const classRenames = renameClashes(doc.context, root, donor.context, donorRoot, 'ClassMap')
  for (const element of donorElements) renameRoleAndClasses(element, roleRenames, classRenames)
  renameClashingIds(elementsUnder(doc.context, root), donorElements, donor.context)
  // A kid's /P points to the donor's root or /Document, which would bring every donor page along.
  for (const kid of kids) {
    const element = donor.context.lookup(kid)
    if (element instanceof PDFDict) element.delete(P)
  }

  const Lang = PDFName.of('Lang')
  const donorLang = (unwrap ? donorDocument.element.lookup(Lang) : undefined) ?? donor.catalog.lookup(Lang)
  const ownLang = ownDocument?.element.lookup(Lang) ?? doc.catalog.lookup(Lang)
  const copiedKids = kids.map(kid => {
    const copy = copier.copy(kid)
    const element = doc.context.lookup(copy)
    if (!(element instanceof PDFDict)) return copy
    element.set(P, into.ref)
    // Each file names its own language; the pages from another file keep theirs.
    if (donorLang !== undefined && donorLang.toString() !== ownLang?.toString() && !element.has(Lang)) element.set(Lang, donorLang.clone())
    return copy
  })
  setKids(into.element, [...kidsOf(into.element), ...copiedKids])

  mergeMap(doc.context, root, donorRoot, 'RoleMap', roleRenames, copier)
  mergeMap(doc.context, root, donorRoot, 'ClassMap', classRenames, copier)
  return true
}

/** Every element under `root`, each once. */
function elementsUnder(context: PDFContext, root: PDFDict): PDFDict[] {
  const found: PDFDict[] = []
  const visited = new Set<PDFDict>([root])
  const pending = [root]
  for (let node = pending.pop(); node !== undefined; node = pending.pop()) {
    for (const kid of kidsOf(node)) {
      const read = readKid(context, kid, undefined)
      if (read?.kind !== 'element' || visited.has(read.dict)) continue
      visited.add(read.dict)
      found.push(read.dict)
      pending.push(read.dict)
    }
  }
  return found
}

/** `name_2`, `name_3`, … : the first that `taken` says is free. */
function freeName(name: string, taken: (candidate: string) => boolean): string {
  let copy = 2
  while (taken(`${name}_${copy}`)) copy += 1
  return `${name}_${copy}`
}

/**
 * The donor's role (RoleMap) or class (ClassMap) names that this file defines differently,
 * each with a free name to use instead, so each file's elements keep their meaning when the
 * two maps are merged: two files can map /Custom to /P and to /H1.
 */
function renameClashes(context: PDFContext, root: PDFDict, donorContext: PDFContext, donorRoot: PDFDict, key: string): Map<string, string> {
  const renames = new Map<string, string>()
  const ours = root.lookupMaybe(PDFName.of(key), PDFDict)
  const theirs = donorRoot.lookupMaybe(PDFName.of(key), PDFDict)
  if (ours === undefined || theirs === undefined) return renames
  for (const [name, value] of theirs.entries()) {
    const mine = ours.get(name)
    if (mine === undefined || context.lookup(mine)?.toString() === donorContext.lookup(value)?.toString()) continue
    renames.set(name.decodeText(), freeName(name.decodeText(), candidate => ours.has(PDFName.of(candidate)) || theirs.has(PDFName.of(candidate))))
  }
  return renames
}

function renamed(name: PDFObject | undefined, renames: ReadonlyMap<string, string>): PDFName | undefined {
  const replacement = name instanceof PDFName ? renames.get(name.decodeText()) : undefined
  return replacement === undefined ? undefined : PDFName.of(replacement)
}

function renameRoleAndClasses(element: PDFDict, roles: ReadonlyMap<string, string>, classes: ReadonlyMap<string, string>): void {
  const role = renamed(element.get(S), roles)
  if (role !== undefined) element.set(S, role)
  const C = PDFName.of('C')
  const own = element.get(C)
  const list = element.context.lookup(own)
  if (list instanceof PDFArray) list.asArray().forEach((entry, index) => {
    const name = renamed(entry, classes)
    if (name !== undefined) list.set(index, name)
  })
  const name = renamed(own, classes)
  if (name !== undefined) element.set(C, name)
}

/** Adds the donor's map entries under their new names; a renamed role mapped onto another renamed role follows it. */
function mergeMap(context: PDFContext, root: PDFDict, donorRoot: PDFDict, key: string, renames: ReadonlyMap<string, string>, copier: PDFObjectCopier): void {
  const theirs = donorRoot.lookupMaybe(PDFName.of(key), PDFDict)
  if (theirs === undefined) return
  let ours = root.lookupMaybe(PDFName.of(key), PDFDict)
  if (ours === undefined) {
    ours = context.obj({})
    root.set(PDFName.of(key), ours)
  }
  for (const [name, value] of theirs.entries()) {
    const target = renamed(name, renames) ?? name
    if (!ours.has(target)) ours.set(target, renamed(value, renames) ?? copier.copy(value))
  }
}

const ID = PDFName.of('ID')

/** A byte string's bytes, one character each, so that strings compare and sort as their bytes do. */
function byteKey(object: PDFObject | undefined): string | undefined {
  if (!(object instanceof PDFString || object instanceof PDFHexString)) return undefined
  return Array.from(object.asBytes(), byte => String.fromCharCode(byte)).join('')
}

function byteString(key: string): PDFHexString {
  return PDFHexString.of(Array.from(key, character => character.charCodeAt(0).toString(16).padStart(2, '0')).join(''))
}

/**
 * Gives the donor's element IDs that this file already uses new ones, and updates the table
 * cells that name them in /Headers: otherwise a donor cell could be read with this file's header.
 */
function renameClashingIds(own: readonly PDFDict[], donorElements: readonly PDFDict[], donorContext: PDFContext): void {
  const taken = new Set(own.flatMap(element => byteKey(element.lookup(ID)) ?? []))
  const renames = new Map<string, string>()
  for (const element of donorElements) {
    const id = byteKey(element.lookup(ID))
    if (id === undefined) continue
    if (taken.has(id)) {
      const fresh = freeName(id, candidate => taken.has(candidate))
      renames.set(id, fresh)
      element.set(ID, byteString(fresh))
      taken.add(fresh)
    } else {
      taken.add(id)
    }
  }
  if (renames.size === 0) return
  for (const element of donorElements) {
    const attributes = donorContext.lookup(element.get(PDFName.of('A')))
    for (const entry of attributes instanceof PDFArray ? attributes.asArray() : [attributes]) {
      const headers = donorContext.lookup(entry)
      const list = headers instanceof PDFDict ? headers.lookup(PDFName.of('Headers')) : undefined
      if (!(list instanceof PDFArray)) continue
      list.asArray().forEach((header, index) => {
        const fresh = renames.get(byteKey(donorContext.lookup(header)) ?? '')
        if (fresh !== undefined) list.set(index, byteString(fresh))
      })
    }
  }
}

/** Marks the document as tagged, for when every page in it came from a tagged file. */
export function markTagged(doc: PDFDocument): void {
  doc.catalog.set(PDFName.of('MarkInfo'), doc.context.obj({ Marked: true }))
}

/** The highest MCID a page's content stream sets, or -1. A page whose content can't be read gives -1. */
function highestMcidInContent(doc: PDFDocument, page: PDFPage): number {
  const contents = doc.context.lookup(page.node.get(PDFName.of('Contents')))
  const streams = contents instanceof PDFArray ? contents.asArray() : contents === undefined ? [] : [contents]
  let highest = -1
  for (const entry of streams) {
    const stream = doc.context.lookup(entry)
    let bytes: Uint8Array | undefined
    try {
      if (stream instanceof PDFContentStream) bytes = stream.getUnencodedContents()
      else if (stream instanceof PDFRawStream) bytes = decodePDFRawStream(stream).decode()
    } catch {
      bytes = undefined
    }
    if (bytes === undefined) continue
    for (const match of new TextDecoder('latin1').decode(bytes).matchAll(/\/MCID\s+(\d+)/g)) highest = Math.max(highest, Number(match[1]))
  }
  return highest
}

export interface ContentTagger {
  /**
   * Starts tagging content about to be drawn on `page`: adds an element of `role` to the tree
   * at the end of what the page shows, and returns the operator that opens its marked content.
   * The caller closes it with `endMarkedContent()`.
   */
  begin(page: PDFPage, role: string, alt?: string): PDFOperator
  /** Turns an annotation's /OBJR into marked content on its page, for a flattened form field. Returns the open operator, or undefined when the annotation has no tag. */
  replaceObject(page: PDFPage, annotation: PDFRef): PDFOperator | undefined
}

/** `/Role <</MCID n>> BDC`. A string operand is written to the content stream as it is. */
function markedContent(role: string, mcid: number): PDFOperator {
  return PDFOperator.of(PDFOperatorNames.BeginMarkedContentSequence, [PDFName.of(role), `<</MCID ${mcid}>>`])
}

/**
 * Tags new content in a tagged document, or returns undefined for an untagged one. New
 * content goes after the last element of its page: inside the page's section when the tree
 * groups content by section (a /Part per page, as Acrobat's form tagging does), else next to it.
 */
export function contentTagger(doc: PDFDocument): ContentTagger | undefined {
  const root = rootOf(doc)
  if (root === undefined) return undefined
  const { context } = doc
  const nextMcid = new Map<string, number>()
  const position = new Map(doc.getPages().map((page, index) => [page.ref.tag, index]))

  const takeMcid = (page: PDFPage): number => {
    let next = nextMcid.get(page.ref.tag)
    if (next === undefined) {
      let highest = highestMcidInContent(doc, page)
      const visit = (element: PDFDict, inherited: string | undefined, visited: Set<PDFDict>): void => {
        if (visited.has(element)) return
        visited.add(element)
        const own = pageOf(element, inherited)
        for (const kid of kidsOf(element)) {
          const read = readKid(context, kid, own)
          if (read?.kind === 'element') visit(read.dict, own, visited)
          else if (read?.kind === 'mcid' && read.stream === undefined && read.page === page.ref.tag) highest = Math.max(highest, read.mcid)
        }
      }
      visit(root, undefined, new Set())
      next = highest + 1
    }
    nextMcid.set(page.ref.tag, next + 1)
    return next
  }

  /** Inserts `element` (registered as `ref`) where the content of `page` ends in reading order. */
  const insert = (ref: PDFRef, element: PDFDict, page: PDFPage): void => {
    const tag = page.ref.tag
    const target = position.get(tag) ?? Infinity
    let parent = root
    let parentRef: PDFRef = rootRef(doc, root)
    let parentPage: string | undefined
    for (let depth = 0; depth < 64; depth += 1) {
      const kids = kidsOf(parent)
      const elements = kids.map(kid => readKid(context, kid, parentPage))
      const pages = elements.map(read => read?.kind === 'element' ? contentPages(context, read.dict, parentPage) : undefined)
      let last = -1
      pages.forEach((set, index) => {
        if (set?.has(tag) === true) last = index
      })
      if (last === -1) {
        // Nothing on this page yet at this level: after the last kid that starts on an earlier page.
        pages.forEach((set, index) => {
          if (set !== undefined && Math.min(...[...set].map(page => position.get(page) ?? Infinity)) < target) last = index
        })
      } else {
        const read = elements[last]
        const kid = kids[last]
        if (read?.kind === 'element' && kid instanceof PDFRef && groupingRoles.has(standardRole(root, read.dict) ?? '')) {
          parentPage = pageOf(read.dict, parentPage)
          parent = read.dict
          parentRef = kid
          continue
        }
      }
      kids.splice(last + 1, 0, ref)
      setKids(parent, kids)
      element.set(P, parentRef)
      return
    }
  }

  return {
    begin(page, role, alt) {
      const mcid = takeMcid(page)
      const element = context.obj({ Type: 'StructElem', S: role, Pg: page.ref, K: [mcid] })
      if (alt !== undefined) element.set(PDFName.of('Alt'), PDFHexString.fromText(alt))
      insert(context.register(element), element, page)
      return markedContent(role, mcid)
    },
    replaceObject(page, annotation) {
      const visited = new Set<PDFDict>()
      const find = (element: PDFDict): PDFOperator | undefined => {
        if (visited.has(element)) return undefined
        visited.add(element)
        const kids = kidsOf(element)
        for (const [index, kid] of kids.entries()) {
          const read = readKid(context, kid, undefined)
          if (read?.kind === 'element') {
            const found = find(read.dict)
            if (found !== undefined) return found
          } else if (read?.kind === 'object' && read.target?.tag === annotation.tag) {
            const mcid = takeMcid(page)
            kids[index] = context.obj({ Type: 'MCR', Pg: page.ref, MCID: mcid })
            setKids(element, kids)
            const role = element.lookup(S)
            return markedContent(role instanceof PDFName ? role.decodeText() : 'Span', mcid)
          }
        }
        return undefined
      }
      return find(root)
    },
  }
}

/** Opens marked content that tells screen readers to skip what follows (decoration). */
export function beginArtifact(): PDFOperator {
  return beginMarkedContent('Artifact')
}

/**
 * Rebuilds what the tree implies after pages were moved, removed, copied in or edited:
 * every element's /P, the /ParentTree with fresh /StructParents keys, and the /IDTree.
 * The old parent tree and ID tree are replaced, so entries for removed elements, and the
 * text those elements quote, leave the file.
 */
export function finishStructure(doc: PDFDocument): void {
  const root = rootOf(doc)
  if (root === undefined) return
  const { context } = doc
  pruneStructure(doc, liveContent(doc))

  /** Element by MCID, for each page or form XObject (by reference tag) that holds marked content. */
  const marked = new Map<string, (PDFRef | undefined)[]>()
  const streams = new Map<string, PDFRef>()
  /** Element of each tagged annotation or object, by the object's reference tag. */
  const objects = new Map<string, { readonly target: PDFRef; readonly element: PDFRef }>()
  const ids: [string, PDFObject, PDFRef][] = []
  const visited = new Set<PDFDict>()

  const visit = (element: PDFDict, ref: PDFRef, inherited: string | undefined): void => {
    if (visited.has(element)) return
    visited.add(element)
    const page = pageOf(element, inherited)
    const id = element.get(ID)
    const key = byteKey(element.lookup(ID))
    if (id !== undefined && key !== undefined && element !== root) ids.push([key, id, ref])
    const kids = kidsOf(element)
    let registered = false
    kids.forEach((kid, index) => {
      const read = readKid(context, kid, page)
      if (read?.kind === 'element') {
        // The parent tree must point to elements by reference.
        let kidRef = kid instanceof PDFRef ? kid : undefined
        if (kidRef === undefined) {
          kidRef = context.register(read.dict)
          kids[index] = kidRef
          registered = true
        }
        read.dict.set(P, ref)
        visit(read.dict, kidRef, page)
      } else if (read?.kind === 'mcid') {
        const holder = read.stream?.tag ?? read.page
        if (holder === undefined || read.mcid > highestMcid) return
        const list = marked.get(holder) ?? []
        list[read.mcid] = ref
        marked.set(holder, list)
        if (read.stream !== undefined) streams.set(read.stream.tag, read.stream)
      } else if (read?.kind === 'object' && read.target !== undefined) {
        objects.set(read.target.tag, { target: read.target, element: ref })
      }
    })
    if (registered) setKids(element, kids)
  }
  visit(root, rootRef(doc, root), undefined)

  const nums: PDFObject[] = []
  const entry = (holder: PDFDict, name: PDFName, value: PDFObject): void => {
    const key = nums.length / 2
    holder.set(name, PDFNumber.of(key))
    nums.push(PDFNumber.of(key), value)
  }
  const toArray = (list: readonly (PDFRef | undefined)[]): PDFArray => context.obj(Array.from(list, element => element ?? PDFNull))
  for (const page of doc.getPages()) {
    const list = marked.get(page.ref.tag)
    if (list === undefined) page.node.delete(StructParents)
    else entry(page.node, StructParents, toArray(list))
    for (const annotation of page.node.Annots()?.asArray() ?? []) {
      const dict = context.lookup(annotation)
      if (dict instanceof PDFDict && !(annotation instanceof PDFRef && objects.has(annotation.tag))) dict.delete(StructParent)
    }
  }
  // Content of form XObjects, tagged through an /MCR's /Stm.
  for (const [tag, ref] of streams) {
    const stream = context.lookup(ref)
    const list = marked.get(tag)
    if (stream instanceof PDFStream && list !== undefined) entry(stream.dict, StructParents, toArray(list))
  }
  for (const { target, element } of objects.values()) {
    const object = context.lookup(target)
    const holder = object instanceof PDFStream ? object.dict : object
    if (holder instanceof PDFDict) entry(holder, StructParent, element)
  }

  root.set(PDFName.of('ParentTree'), context.register(context.obj({ Nums: nums })))
  root.set(PDFName.of('ParentTreeNextKey'), PDFNumber.of(nums.length / 2))
  if (ids.length === 0) {
    root.delete(PDFName.of('IDTree'))
  } else {
    // A name tree's keys must be in sorted order.
    ids.sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
    root.set(PDFName.of('IDTree'), context.register(context.obj({ Names: ids.flatMap(([, id, ref]) => [id, ref]) })))
  }
}
