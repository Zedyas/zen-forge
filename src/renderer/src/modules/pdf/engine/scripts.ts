import {
  decodePDFRawStream,
  PDFArray,
  PDFDict,
  PDFHexString,
  PDFName,
  PDFRawStream,
  PDFString,
  type PDFContext,
  type PDFDocument,
  type PDFObject,
  type PDFObjectCopier,
} from '@cantoo/pdf-lib'

interface Script {
  readonly name: string
  /** The JavaScript action dictionary, or a reference to it. */
  readonly action: PDFObject
}

function textOf(object: PDFObject | undefined): string | undefined {
  if (object instanceof PDFString || object instanceof PDFHexString) return object.decodeText()
  if (object instanceof PDFRawStream) return new TextDecoder('latin1').decode(decodePDFRawStream(object).decode())
  return undefined
}

/** The leaves of a name tree, in tree order. A name tree is a /Names array of name–value pairs, or /Kids of smaller trees. */
function nameTreeEntries(context: PDFContext, tree: PDFDict, visited = new Set<PDFDict>()): Script[] {
  if (visited.has(tree)) return []
  visited.add(tree)
  const names = tree.lookupMaybe(PDFName.of('Names'), PDFArray)?.asArray() ?? []
  const leaves: Script[] = []
  for (let index = 0; index + 1 < names.length; index += 2) {
    const name = textOf(context.lookup(names[index]))
    if (name !== undefined) leaves.push({ name, action: names[index + 1] })
  }
  const kids = tree.lookupMaybe(PDFName.of('Kids'), PDFArray)?.asArray() ?? []
  for (const kid of kids) {
    const dict = context.lookup(kid)
    if (dict instanceof PDFDict) leaves.push(...nameTreeEntries(context, dict, visited))
  }
  return leaves
}

/** The document-level scripts: the /JavaScript name tree in the catalog's /Names. */
function documentScripts(doc: PDFDocument): Script[] {
  const tree = doc.catalog.lookupMaybe(PDFName.of('Names'), PDFDict)?.lookupMaybe(PDFName.of('JavaScript'), PDFDict)
  return tree === undefined ? [] : nameTreeEntries(doc.context, tree)
}

function scriptText(context: PDFContext, action: PDFObject): string | undefined {
  const dict = context.lookup(action)
  return dict instanceof PDFDict ? textOf(context.lookup(dict.get(PDFName.of('JS')))) : undefined
}

/**
 * Copies `donor`'s document-level scripts into `doc`, for pages combined from another PDF.
 * A script with the same name and code as one `doc` has, or as a renamed copy, is skipped: forms from the same
 * publisher carry the same boilerplate. A script with a taken name but other code is kept under
 * a new name, because scripts are run by name order and none of them should replace another.
 * Zen Suzu never runs these scripts; they are kept for viewers that do.
 */
export function copyDocumentScripts(doc: PDFDocument, donor: PDFDocument, copier: PDFObjectCopier): void {
  const incoming = documentScripts(donor)
  if (incoming.length === 0) return
  const merged = documentScripts(doc)
  const texts = new Map(merged.map(({ name, action }) => [name, scriptText(doc.context, action)]))
  for (const { name, action } of incoming) {
    const text = scriptText(donor.context, action)
    // The script, or an earlier renamed copy of it, may already be there.
    let unique = name
    for (let copy = 2; texts.has(unique) && texts.get(unique) !== text; copy += 1) unique = `${name} (${copy})`
    if (texts.has(unique)) continue
    texts.set(unique, text)
    merged.push({ name: unique, action: copier.copy(action) })
  }

  // A name tree's keys must be in sorted order. One flat /Names array is a valid tree.
  merged.sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0))
  const names = doc.context.obj([])
  for (const { name, action } of merged) {
    names.push(PDFHexString.fromText(name))
    names.push(action)
  }
  let catalogNames = doc.catalog.lookupMaybe(PDFName.of('Names'), PDFDict)
  if (catalogNames === undefined) {
    catalogNames = doc.context.obj({})
    doc.catalog.set(PDFName.of('Names'), doc.context.register(catalogNames))
  }
  catalogNames.set(PDFName.of('JavaScript'), doc.context.register(doc.context.obj({ Names: names })))
}
