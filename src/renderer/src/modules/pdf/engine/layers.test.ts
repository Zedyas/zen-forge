import { decodePDFRawStream, PDFDict, PDFDocument, PDFName, PDFRawStream, PDFString, type PDFPage } from '@cantoo/pdf-lib'
import { describe, expect, it } from 'vitest'
import { scanContent } from './content-stream'
import { removeHiddenLayers } from './layers'
import { extractPageTexts, extractPlacedText, fileContainsText, layeredPdf } from './test-support'

/** A page's content after removal, as text, and the operators in it. */
function pageContent(page: PDFPage): { readonly text: string; readonly operators: string[] } {
  const stream = page.node.lookup(PDFName.of('Contents'))
  if (!(stream instanceof PDFRawStream)) throw new Error('The page content is not one stream')
  const bytes = decodePDFRawStream(stream).decode()
  return { text: new TextDecoder('latin1').decode(bytes), operators: scanContent(bytes).map(instruction => instruction.operator) }
}

describe('removeHiddenLayers', () => {
  it('removes a hidden block whole when nothing after it depends on it', async () => {
    const { doc, page } = await layeredPdf('/OC /hidden BDC q 1 0 0 rg BT /F1 14 Tf 40 500 Td (HIDDEN) Tj ET Q EMC BT /F1 14 Tf 40 300 Td (SHOWN) Tj ET')
    removeHiddenLayers(doc)
    expect(pageContent(page).operators).toEqual(['BT', 'Tf', 'Td', 'Tj', 'ET'])
  })

  it('keeps the colour, position and clip a hidden block leaves behind, and removes what it paints', async () => {
    const { doc, page } = await layeredPdf([
      '/OC /hidden BDC 0 0 1 rg 1 0 0 1 0 -20 cm 0 0 400 400 re W n 10 10 50 50 re f',
      '/Span <</ActualText (HIDDEN-ALT)>> BDC BT /F1 14 Tf 40 500 Td (HIDDEN-TEXT) Tj ET EMC EMC',
      'BT 40 300 Td (SHOWN) Tj ET',
    ].join('\n'))
    removeHiddenLayers(doc)

    const { text, operators } = pageContent(page)
    expect(operators).toEqual(['BMC', 'rg', 'cm', 're', 'W', 'n', 're', 'n', 'BMC', 'BT', 'Tf', 'Td', 'ET', 'EMC', 'EMC', 'BT', 'Td', 'Tj', 'ET'])
    expect(text).not.toContain('HIDDEN')
  })

  it('keeps the lines after hidden lines of a text object where they were', async () => {
    const { doc, page } = await layeredPdf('BT /F1 14 Tf 16 TL 40 500 Td /OC /hidden BDC (HIDDEN-1) \' 2 1 (HIDDEN-2) " EMC (SHOWN) \' ET')
    const shown = async (): Promise<unknown> => (await extractPlacedText(await doc.save(), 0)).find(item => item.text === 'SHOWN')
    const before = await shown()
    removeHiddenLayers(doc)

    expect(await shown()).toEqual(before)
    expect(pageContent(page).text).not.toContain('HIDDEN')
  })

  it('stops rather than move visible text shown after hidden text on the same line', async () => {
    const { doc } = await layeredPdf('BT /F1 12 Tf 72 700 Td /OC /hidden BDC (DRAFT ) Tj EMC (Final) Tj ET')
    expect(() => removeHiddenLayers(doc)).toThrow('mixes a hidden layer into visible text')
  })

  it('rewrites a visible form XObject that holds hidden content', async () => {
    const { doc, page, layers } = await layeredPdf('/Inner Do', {
      Inner: { content: '/OC /hidden BDC BT /F1 14 Tf 40 500 Td (HIDDEN-IN-FORM) Tj ET EMC BT /F1 14 Tf 40 300 Td (SHOWN-IN-FORM) Tj ET' },
    })
    const inner = doc.context.lookup(page.node.Resources()?.lookup(PDFName.of('XObject'), PDFDict).get(PDFName.of('Inner')))
    if (!(inner instanceof PDFRawStream)) throw new Error('The form is missing')
    inner.dict.lookup(PDFName.of('Resources'), PDFDict).set(PDFName.of('Properties'), doc.context.obj({ hidden: layers.hidden }))
    removeHiddenLayers(doc)

    const bytes = await doc.save()
    expect(await fileContainsText(bytes, 'HIDDEN-IN-FORM')).toBe(false)
    expect(await extractPageTexts(bytes)).toEqual(['SHOWN-IN-FORM'])
  })

  it('removes a form field whose widgets are all hidden, value included', async () => {
    const { doc, page, layers } = await layeredPdf('')
    const field = doc.getForm().createTextField('ssn')
    field.setText('SECRET-FIELD-VALUE')
    field.addToPage(page, { x: 40, y: 400, width: 200, height: 20 })
    field.acroField.getWidgets()[0]?.dict.set(PDFName.of('OC'), layers.hidden)
    // Loaded from bytes, as when saving: pdf-lib then wraps the content of a page it removes an annotation from.
    const loaded = await PDFDocument.load(await doc.save())
    removeHiddenLayers(loaded)

    const bytes = await loaded.save()
    expect(await fileContainsText(bytes, 'SECRET-FIELD-VALUE')).toBe(false)
    expect((await PDFDocument.load(bytes)).getForm().getFields()).toEqual([])
  })

  it('removes hidden content from annotation appearances, tiling patterns and Type 3 glyphs', async () => {
    const { doc, page, layers } = await layeredPdf('')
    const { context } = doc
    const resources = page.node.Resources()
    const font = resources?.lookup(PDFName.of('Font'), PDFDict).get(PDFName.of('F1'))
    const own = { Font: { F1: font }, Properties: { hidden: layers.hidden } }
    const hiddenText = (marker: string): string => `/OC /hidden BDC BT /F1 10 Tf 2 5 Td (${marker}) Tj ET EMC`

    const appearance = context.register(context.stream(hiddenText('HIDDEN-APPEARANCE'), { Type: 'XObject', Subtype: 'Form', BBox: [0, 0, 200, 20], Resources: own }))
    page.node.addAnnot(context.register(context.obj({ Type: 'Annot', Subtype: 'Square', Rect: [40, 400, 240, 420], AP: { N: appearance } })))
    const pattern = context.register(context.stream(hiddenText('HIDDEN-PATTERN'), { PatternType: 1, PaintType: 1, TilingType: 1, BBox: [0, 0, 100, 100], XStep: 100, YStep: 100, Resources: own }))
    const glyph = context.register(context.stream(`1000 0 d0 ${hiddenText('HIDDEN-GLYPH')}`))
    const type3 = context.register(context.obj({
      Type: 'Font', Subtype: 'Type3', FontBBox: [0, 0, 1000, 1000], FontMatrix: [0.001, 0, 0, 0.001, 0, 0],
      CharProcs: { a: glyph }, Encoding: { Type: 'Encoding', Differences: [97, 'a'] }, FirstChar: 97, LastChar: 97, Widths: [1000], Resources: own,
    }))
    resources?.set(PDFName.of('Pattern'), context.obj({ P1: pattern }))
    resources?.lookup(PDFName.of('Font'), PDFDict).set(PDFName.of('T3'), type3)
    removeHiddenLayers(doc)

    const bytes = await doc.save()
    for (const marker of ['HIDDEN-APPEARANCE', 'HIDDEN-PATTERN', 'HIDDEN-GLYPH']) expect(await fileContainsText(bytes, marker)).toBe(false)
  })

  it('stops when something it does not clean still points to a layer, which would show once the layers are gone', async () => {
    const { doc, page, layers } = await layeredPdf('')
    const { context } = doc
    const turnOff = context.obj({ S: 'SetOCGState', State: ['OFF', layers.shown] })
    const link = context.obj({ Type: 'Annot', Subtype: 'Link', Rect: [40, 40, 120, 60], A: { S: 'URI', URI: PDFString.of('https://example.com'), Next: turnOff } })
    page.node.addAnnot(context.register(link))
    expect(() => removeHiddenLayers(doc)).toThrow('hidden layer in a place Zendo can’t clean')
  })
})
