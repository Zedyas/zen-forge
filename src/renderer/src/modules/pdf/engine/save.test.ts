import { PDFDocument, StandardFonts, degrees } from '@cantoo/pdf-lib'
import { describe, expect, it } from 'vitest'
import { savePdf, extractPages } from './save'
import { extractPageTexts, extractPlacedText, fileContainsText } from './test-support'

async function documentWith(
  labels: readonly string[],
  rotation = 0,
  size: readonly [number, number] = [400, 600],
): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  for (const label of labels) {
    const page = doc.addPage([size[0], size[1]])
    if (rotation !== 0) page.setRotation(degrees(rotation))
    page.drawText(label, { x: 40, y: 500, size: 18, font })
  }
  return doc.save()
}

describe('savePdf page assembly', () => {
  it('reorders, deletes, rotates and merges in one pass', async () => {
    const primary = await documentWith(['ALPHA', 'BRAVO', 'CHARLIE'])
    const secondary = await documentWith(['ZULU'])

    const output = await savePdf({
      sources: [primary, secondary],
      pages: [
        { source: 0, index: 2, rotation: 90, edits: [] },
        { source: 1, index: 0, rotation: 0, edits: [] },
        { source: 0, index: 0, rotation: 180, edits: [] },
      ],
    })

    expect(await extractPageTexts(output)).toEqual(['CHARLIE', 'ZULU', 'ALPHA'])

    const reopened = await PDFDocument.load(output)
    expect(reopened.getPageCount()).toBe(3)
    expect(reopened.getPages().map(page => page.getRotation().angle)).toEqual([90, 0, 180])
  })

  it('adds the user rotation on top of the page rotation', async () => {
    const primary = await documentWith(['ALPHA'], 270)

    const output = await savePdf({
      sources: [primary],
      pages: [{ source: 0, index: 0, rotation: 180, edits: [] }],
    })

    const reopened = await PDFDocument.load(output)
    expect(reopened.getPages()[0].getRotation().angle).toBe(90)
  })

  it('refuses a source page listed twice, since a page object can sit in the page tree only once', async () => {
    const primary = await documentWith(['ALPHA', 'BRAVO'])
    const twice = { source: 0, index: 1, rotation: 0, edits: [] } as const

    await expect(savePdf({ sources: [primary], pages: [twice, twice] })).rejects.toThrow('listed twice')
  })
})

describe('deleted content', () => {
  it('leaves nothing of a deleted page in the saved file', async () => {
    const primary = await documentWith(['ALPHA', 'SALARYTABLE'])
    expect(await fileContainsText(primary, 'SALARYTABLE')).toBe(true)

    const output = await savePdf({ sources: [primary], pages: [{ source: 0, index: 0, rotation: 0, edits: [] }] })

    expect(await extractPageTexts(output)).toEqual(['ALPHA'])
    expect(await fileContainsText(output, 'SALARYTABLE')).toBe(false)
  })
})

describe('text edits', () => {
  it('places text where the viewer sees it on a rotated page', async () => {
    const primary = await documentWith([''], 90)
    const metrics = await PDFDocument.create()
    const ascent = (await metrics.embedFont(StandardFonts.Helvetica)).heightAtSize(12, {
      descender: false,
    })

    const output = await savePdf({
      sources: [primary],
      pages: [
        {
          source: 0,
          index: 0,
          rotation: 0,
          edits: [{ kind: 'text', x: 100, y: 50, text: 'EDGE', size: 12, color: '#102030' }],
        },
      ],
    })

    const placed = await extractPlacedText(output, 0)
    const edge = placed.find(item => item.text === 'EDGE')
    expect(edge).toBeDefined()
    expect(edge?.x).toBeCloseTo(100, 1)
    expect(edge?.y).toBeCloseTo(50 + ascent, 1)
  })
})

describe('extractPages', () => {
  it('keeps only the requested pages, in the requested order', async () => {
    const source = await documentWith(['ALPHA', 'BRAVO', 'CHARLIE'])
    const output = await extractPages(source, [2, 0])
    expect(await extractPageTexts(output)).toEqual(['CHARLIE', 'ALPHA'])
  })
})
