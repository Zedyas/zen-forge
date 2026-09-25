import { PDFDict, PDFDocument, PDFName, PDFString, StandardFonts, degrees } from '@cantoo/pdf-lib'
import { describe, expect, it } from 'vitest'
import { pageGeometry, rectToDisplayed } from './geometry'
import { redactPages } from './redact'
import { extractPageTexts, fileContainsText, layeredPdf } from './test-support'

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

  it('removes covered text in a hidden layer, marked in the page content or on a form XObject', async () => {
    // The box covers the line at y = 500 (82pt from the top as displayed).
    const box = { x: 36, y: 82, width: 170, height: 24 }
    const { doc } = await layeredPdf(
      '/OC /hidden BDC BT /F1 14 Tf 40 500 Td (HIDDEN-MARKED) Tj ET EMC /Layer Do BT /F1 14 Tf 40 300 Td (KEEP ME) Tj ET',
      { Layer: { content: 'BT /F1 14 Tf 40 505 Td (HIDDEN-FORM) Tj ET', layer: 'hidden' } },
    )
    const bytes = await doc.save({ useObjectStreams: false })
    expect(await fileContainsText(bytes, 'HIDDEN-MARKED')).toBe(true)

    const output = await redactPages(bytes, [{ index: 0, boxes: [box] }], keepHidden)
    expect(await fileContainsText(output, 'HIDDEN-MARKED')).toBe(false)
    expect(await fileContainsText(output, 'HIDDEN-FORM')).toBe(false)
    const [text] = await extractPageTexts(output)
    expect(text).toContain('KEEP ME')
  })
})

describe('hidden layers', () => {
  // The box sits on a second, blank page, so the layered page reaches hidden-layer removal as written.
  const boxOnBlankPage = [{ index: 1, boxes: [{ x: 10, y: 10, width: 20, height: 20 }] }]

  it('removes hidden layers with the hidden information and leaves the rest always visible', async () => {
    // Unfiltered image data holding `EI (`, which a scanner reading it as tokens would take for the image's end and a string.
    const inlineImage = 'q 4 0 0 1 300 50 cm BI /W 4 /H 1 /CS /G /BPC 8 ID EI (\nEI Q'
    const { doc, page, layers } = await layeredPdf([
      inlineImage,
      '/OC /hidden BDC BT /F1 14 Tf 40 500 Td (HIDDEN-TEXT) Tj ET',
      '/Span <</MCID 0>> BDC BT /F1 14 Tf 40 480 Td (HIDDEN-NESTED) Tj ET EMC',
      'BT /F1 14 Tf 40 460 Td (HIDDEN-AFTER-NESTED) Tj ET /Inner Do EMC',
      '/OC /shown BDC BT /F1 14 Tf 40 300 Td (SHOWN-TEXT) Tj ET EMC',
      '/HiddenForm Do',
    ].join('\n'), {
      Inner: { content: 'BT /F1 14 Tf 40 100 Td (HIDDEN-INNER) Tj ET' },
      HiddenForm: { content: 'BT /F1 14 Tf 40 200 Td (HIDDEN-FORM) Tj ET', layer: 'hidden' },
    })
    const { context } = doc
    const link = context.obj({ Type: 'Annot', Subtype: 'Link', Rect: [40, 40, 120, 60], OC: layers.hidden, A: { S: 'URI', URI: PDFString.of('https://HIDDEN-LINK') } })
    page.node.addAnnot(context.register(link))
    doc.addPage([400, 600])
    const bytes = await doc.save()
    const markers = ['HIDDEN-TEXT', 'HIDDEN-NESTED', 'HIDDEN-AFTER-NESTED', 'HIDDEN-INNER', 'HIDDEN-FORM', 'HIDDEN-LINK', 'Hidden layer', 'Shown layer']

    const kept = await redactPages(bytes, boxOnBlankPage, keepHidden)
    for (const marker of markers) expect(await fileContainsText(kept, marker)).toBe(true)

    const cleaned = await redactPages(bytes, boxOnBlankPage, removeHidden)
    const left: string[] = []
    for (const marker of markers) if (await fileContainsText(cleaned, marker)) left.push(marker)
    expect(left).toEqual([])
    const [text] = await extractPageTexts(cleaned)
    expect(text).toContain('SHOWN-TEXT')
    expect(await fileContainsText(cleaned, 'BI /W 4 /H 1')).toBe(true)
    expect((await PDFDocument.load(cleaned)).catalog.get(PDFName.of('OCProperties'))).toBeUndefined()
  })

  it('reads membership dictionaries: policies over groups, and visibility expressions', async () => {
    const { doc, page, layers: { hidden, shown } } = await layeredPdf('')
    const { context } = doc
    const memberships: Array<[name: string, membership: PDFDict, visible: boolean]> = [
      ['AnyOn', context.obj({ Type: 'OCMD', OCGs: [hidden, shown], P: 'AnyOn' }), true],
      ['AllOn', context.obj({ Type: 'OCMD', OCGs: [hidden, shown], P: 'AllOn' }), false],
      ['AnyOff', context.obj({ Type: 'OCMD', OCGs: [hidden, shown], P: 'AnyOff' }), true],
      ['AllOff', context.obj({ Type: 'OCMD', OCGs: [hidden, shown], P: 'AllOff' }), false],
      ['ShownAndNotHidden', context.obj({ Type: 'OCMD', VE: ['And', shown, ['Not', hidden]] }), true],
      ['HiddenOrNothing', context.obj({ Type: 'OCMD', VE: ['Or', hidden] }), false],
    ]
    const propertyLists = page.node.Resources()?.lookup(PDFName.of('Properties'), PDFDict)
    memberships.forEach(([name, membership]) => propertyLists?.set(PDFName.of(name), context.register(membership)))
    const content = memberships.map(([name], line) => `/OC /${name} BDC BT /F1 10 Tf 40 ${500 - line * 20} Td (MEMBER-${name}) Tj ET EMC`)
    page.node.set(PDFName.of('Contents'), context.register(context.stream(content.join('\n'))))
    doc.addPage([400, 600])

    const output = await redactPages(await doc.save(), boxOnBlankPage, removeHidden)
    const [text] = await extractPageTexts(output)
    expect(memberships.filter(([name]) => text?.includes(`MEMBER-${name}`)).map(([name]) => name)).toEqual(['AnyOn', 'AnyOff', 'ShownAndNotHidden'])
  })
})

