import { PDFDict, PDFDocument, PDFName, PDFString, StandardFonts, degrees } from '@cantoo/pdf-lib'
import { describe, expect, it } from 'vitest'
import { pageGeometry, rectToDisplayed } from './geometry'
import { redactPages } from './redact'
import { extractPageTexts, fileContainsText } from './test-support'

const keepHidden = { removeHiddenInformation: false }
const removeHidden = { removeHiddenInformation: true }

/** One page with a line to redact and a line to keep; the box is in displayed coordinates, as the UI sends it. */
async function statement(rotation = 0) {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const page = doc.addPage([400, 600])
  page.setRotation(degrees(rotation))
  page.drawText('ACCOUNT 4417-2291', { x: 40, y: 500, size: 14, font })
  page.drawText('KEEP ME', { x: 40, y: 300, size: 14, font })
  const box = rectToDisplayed(pageGeometry(page), { x: 36, y: 494, width: 170, height: 24 })
  return { doc, page, box }
}

describe('redactPages', () => {
  it.each([0, 90])('removes only the covered text and keeps the rest of the page as text (rotation %i)', async rotation => {
    const { doc, box } = await statement(rotation)
    const output = await redactPages(await doc.save(), [{ index: 0, boxes: [box] }], keepHidden)

    const [text] = await extractPageTexts(output)
    expect(text).toContain('KEEP ME')
    expect(text).not.toContain('4417')
    expect(await fileContainsText(output, '4417')).toBe(false)

    const page = (await PDFDocument.load(output)).getPages()[0]
    expect(page?.getSize()).toEqual({ width: 400, height: 600 })
    expect(page?.getRotation().angle).toBe(rotation)
  })

  it('removes a form field under a box, value included, and keeps the fields elsewhere', async () => {
    const { doc, page, box } = await statement()
    const form = doc.getForm()
    const iban = form.createTextField('iban')
    iban.setText('IBANSECRET')
    iban.addToPage(page, { x: 220, y: 496, width: 150, height: 20 })
    const name = form.createTextField('name')
    name.setText('KEEPNAME')
    name.addToPage(page, { x: 220, y: 296, width: 150, height: 20 })
    const wide = { ...box, width: 360 }

    const output = await redactPages(await doc.save(), [{ index: 0, boxes: [wide] }], keepHidden)

    expect(await fileContainsText(output, 'IBANSECRET')).toBe(false)
    const fields = (await PDFDocument.load(output)).getForm().getFields().map(field => field.getName())
    expect(fields).toEqual(['name'])
  })

  it('removes copies of the text outside the page content, and hidden information on request', async () => {
    const build = async () => {
      const { doc, page, box } = await statement()
      const { context } = doc
      const secret = (label: string): PDFString => PDFString.of(`SECRET-${label}`)
      // Always removed on a redacted page: a stored thumbnail, application data, tag alt text.
      page.node.set(PDFName.of('Thumb'), context.register(context.stream('BT (SECRET-THUMB) Tj ET')))
      page.node.set(PDFName.of('PieceInfo'), context.obj({ App: { Private: secret('PIECE') } }))
      const element = context.register(context.obj({ Type: 'StructElem', S: 'Figure', Pg: page.ref, Alt: secret('ALT') }))
      doc.catalog.set(PDFName.of('StructTreeRoot'), context.register(context.obj({ Type: 'StructTreeRoot', K: [element] })))
      // Hidden information: metadata, a bookmark, an attachment and a comment.
      context.lookup(context.trailerInfo.Info, PDFDict).set(PDFName.of('Title'), secret('TITLE'))
      doc.catalog.set(PDFName.of('Metadata'), context.register(context.stream('<x:xmpmeta>SECRET-XMP</x:xmpmeta>')))
      const outlines = context.nextRef()
      const bookmark = context.register(context.obj({ Title: secret('BOOKMARK'), Parent: outlines, Dest: [page.ref, 'Fit'] }))
      context.assign(outlines, context.obj({ Type: 'Outlines', First: bookmark, Last: bookmark, Count: 1 }))
      doc.catalog.set(PDFName.of('Outlines'), outlines)
      await doc.attach(new TextEncoder().encode('SECRET-ATTACHMENT'), 'notes.txt')
      page.node.addAnnot(context.register(context.obj({ Type: 'Annot', Subtype: 'Text', Rect: [300, 100, 320, 120], Contents: secret('COMMENT') })))
      return { bytes: await doc.save({ useObjectStreams: false }), box }
    }

    const { bytes, box } = await build()
    expect(await fileContainsText(bytes, 'SECRET-ALT')).toBe(true)

    const kept = await redactPages(bytes, [{ index: 0, boxes: [box] }], keepHidden)
    for (const label of ['THUMB', 'PIECE', 'ALT']) expect(await fileContainsText(kept, `SECRET-${label}`)).toBe(false)
    expect(await fileContainsText(kept, 'SECRET-TITLE')).toBe(true)

    const cleaned = await redactPages(bytes, [{ index: 0, boxes: [box] }], removeHidden)
    const left: string[] = []
    for (const label of ['THUMB', 'PIECE', 'ALT', 'TITLE', 'XMP', 'BOOKMARK', 'ATTACHMENT', 'COMMENT']) if (await fileContainsText(cleaned, `SECRET-${label}`)) left.push(label)
    expect(left).toEqual([])
    expect(await fileContainsText(cleaned, 'SECRET')).toBe(false)
    const [text] = await extractPageTexts(cleaned)
    expect(text).toContain('KEEP ME')
  })
})
