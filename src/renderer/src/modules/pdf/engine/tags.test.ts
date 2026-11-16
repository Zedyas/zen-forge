import {
  decodePDFRawStream,
  endMarkedContent,
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFOperator,
  PDFOperatorNames,
  PDFRawStream,
  PDFRef,
  StandardFonts,
} from '@cantoo/pdf-lib'
import { describe, expect, it } from 'vitest'
import { savePdf } from './save'
import { extractPageTexts } from './test-support'
import type { PageRef } from './types'

/** A tagged PDF with one page per label: /Document holds one /P per page, the label as its marked content (MCID 0). */
async function taggedDocument(labels: readonly string[], { field = false } = {}): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const { context } = doc
  const documentRef = context.nextRef()
  const paragraphs: PDFRef[] = []
  const nums: (number | PDFRef | PDFArray)[] = []
  for (const [index, label] of labels.entries()) {
    const page = doc.addPage([400, 600])
    page.pushOperators(PDFOperator.of(PDFOperatorNames.BeginMarkedContentSequence, [PDFName.of('P'), '<</MCID 0>>']))
    page.drawText(label, { x: 40, y: 500, size: 18, font })
    page.pushOperators(endMarkedContent())
    const paragraph = context.register(context.obj({ Type: 'StructElem', S: 'P', P: documentRef, Pg: page.ref, K: [0] }))
    paragraphs.push(paragraph)
    page.node.set(PDFName.of('StructParents'), PDFNumber.of(index))
    nums.push(index, context.obj([paragraph]))
  }
  const kids: PDFRef[] = [...paragraphs]
  if (field) {
    const page = doc.getPage(0)
    doc.getForm().createTextField('name').addToPage(page, { x: 40, y: 300, width: 200, height: 24 })
    doc.getForm().getTextField('name').setText('Jane')
    const widget = page.node.Annots()?.get(0)
    if (!(widget instanceof PDFRef)) throw new Error('no widget')
    const form = context.register(context.obj({ Type: 'StructElem', S: 'Form', P: documentRef, Pg: page.ref, K: [context.obj({ Type: 'OBJR', Obj: widget, Pg: page.ref })] }))
    kids.push(form)
    context.lookup(widget, PDFDict).set(PDFName.of('StructParent'), PDFNumber.of(labels.length))
    nums.push(labels.length, form)
  }
  context.assign(documentRef, context.obj({ Type: 'StructElem', S: 'Document', K: kids }))
  const root = context.obj({ Type: 'StructTreeRoot', K: [documentRef], ParentTree: context.obj({ Nums: nums }), ParentTreeNextKey: nums.length / 2 })
  doc.catalog.set(PDFName.of('StructTreeRoot'), context.register(root))
  doc.catalog.set(PDFName.of('MarkInfo'), context.obj({ Marked: true }))
  return doc.save()
}

function pageContent(doc: PDFDocument, index: number): string {
  const contents = doc.context.lookup(doc.getPage(index).node.get(PDFName.of('Contents')))
  const streams = contents instanceof PDFArray ? contents.asArray() : [contents]
  return streams.map(stream => {
    const raw = doc.context.lookup(stream)
    return raw instanceof PDFRawStream ? new TextDecoder('latin1').decode(decodePDFRawStream(raw).decode()) : ''
  }).join('\n')
}

function page(source: number, index: number, edits: PageRef['edits'] = []): PageRef {
  return { source, index, rotation: 0, edits }
}

interface Read {
  /** The role of each element under /Document, with the page its content is on (by position). */
  readonly order: string[]
  readonly doc: PDFDocument
  /** The element the parent tree gives for a page's MCID. */
  readonly parentOf: (pageIndex: number, mcid: number) => PDFDict | undefined
}

async function readTags(bytes: Uint8Array): Promise<Read> {
  const doc = await PDFDocument.load(bytes)
  const { context } = doc
  const pages = doc.getPages()
  const root = doc.catalog.lookup(PDFName.of('StructTreeRoot'), PDFDict)
  const document = context.lookup(root.lookup(PDFName.of('K'), PDFArray).get(0), PDFDict)
  const order = document.lookup(PDFName.of('K'), PDFArray).asArray().map(kid => {
    const element = context.lookup(kid, PDFDict)
    return `${element.lookup(PDFName.of('S'), PDFName).decodeText()}@${pages.findIndex(candidate => candidate.ref === element.get(PDFName.of('Pg')))}`
  })
  const nums = context.lookup(root.get(PDFName.of('ParentTree')), PDFDict).lookup(PDFName.of('Nums'), PDFArray).asArray()
  const parentOf = (pageIndex: number, mcid: number): PDFDict | undefined => {
    const key = pages[pageIndex].node.lookup(PDFName.of('StructParents'))
    const at = nums.findIndex((entry, index) => index % 2 === 0 && entry instanceof PDFNumber && key instanceof PDFNumber && entry.asNumber() === key.asNumber())
    const list = at === -1 ? undefined : context.lookup(nums[at + 1])
    const element = list instanceof PDFArray ? context.lookup(list.get(mcid)) : undefined
    return element instanceof PDFDict ? element : undefined
  }
  return { order, doc, parentOf }
}

describe('accessibility tags on save', () => {
  it('follows reordered and removed pages, and points each page to its own tags', async () => {
    const source = await taggedDocument(['ALPHA', 'BRAVO', 'CHARLIE'])

    const output = await savePdf({ sources: [source], pages: [page(0, 2), page(0, 0)] })

    const { order, parentOf } = await readTags(output)
    expect(await extractPageTexts(output)).toEqual(['CHARLIE', 'ALPHA'])
    // BRAVO's paragraph went with its page; the others are read in the new page order.
    expect(order).toEqual(['P@0', 'P@1'])
    expect(parentOf(0, 0)?.get(PDFName.of('Pg'))).toBe((await PDFDocument.load(output)).getPage(0).ref)
  })

  it('keeps the tags of pages combined from another tagged file, in page order', async () => {
    const first = await taggedDocument(['ALPHA', 'BRAVO'])
    const second = await taggedDocument(['YANKEE', 'ZULU'])

    const output = await savePdf({ sources: [first, second], pages: [page(0, 0), page(1, 1), page(0, 1)] })

    const { order, doc, parentOf } = await readTags(output)
    expect(await extractPageTexts(output)).toEqual(['ALPHA', 'ZULU', 'BRAVO'])
    expect(order).toEqual(['P@0', 'P@1', 'P@2'])
    // The copied page's MCID 0 leads to the copied paragraph, not to a paragraph of the first file.
    expect(parentOf(1, 0)?.get(PDFName.of('Pg'))).toBe(doc.getPage(1).ref)
    expect(String(doc.catalog.get(PDFName.of('MarkInfo')))).toContain('/Marked true')
  })

  it('tags added text and signatures after the page’s content, and marks highlights as decoration', async () => {
    const source = await taggedDocument(['ALPHA', 'BRAVO'])
    const png = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='), character => character.charCodeAt(0))

    const output = await savePdf({
      sources: [source],
      pages: [
        page(0, 0, [
          { kind: 'text', x: 40, y: 200, text: 'Note', size: 12, color: '#000000' },
          { kind: 'image', x: 40, y: 300, width: 80, height: 20, png },
          { kind: 'rect', x: 10, y: 10, width: 50, height: 10, fill: '#ffff00', opacity: 0.35 },
        ]),
        page(0, 1),
      ],
    })

    const { order, parentOf } = await readTags(output)
    expect(order).toEqual(['P@0', 'P@0', 'Figure@0', 'P@1'])
    expect(parentOf(0, 1)?.lookup(PDFName.of('S'))).toBe(PDFName.of('P'))
    const alt = parentOf(0, 2)?.lookup(PDFName.of('Alt'))
    expect(alt instanceof PDFHexString ? alt.decodeText() : alt).toBe('Signature')
    expect(pageContent(await PDFDocument.load(output), 0)).toContain('/Artifact BMC')
  })

  it('moves a flattened field’s tag from its widget to the drawn value', async () => {
    const source = await taggedDocument(['ALPHA'], { field: true })

    const output = await savePdf({ sources: [source], pages: [page(0, 0)], flattenForm: true })

    const { order, doc, parentOf } = await readTags(output)
    expect(doc.getPage(0).node.Annots()?.size() ?? 0).toBe(0)
    expect(order).toEqual(['P@0', 'Form@0'])
    expect(parentOf(0, 1)?.lookup(PDFName.of('S'))).toBe(PDFName.of('Form'))
  })

  it('adds no tags to an untagged file', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([400, 600])

    const output = await savePdf({ sources: [await doc.save()], pages: [page(0, 0, [{ kind: 'text', x: 40, y: 200, text: 'Note', size: 12, color: '#000000' }])] })

    expect((await PDFDocument.load(output)).catalog.has(PDFName.of('StructTreeRoot'))).toBe(false)
  })
})
