import { decodePDFRawStream, PDFDict, PDFName, PDFRawStream, type PDFPage } from '@cantoo/pdf-lib'
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
})
