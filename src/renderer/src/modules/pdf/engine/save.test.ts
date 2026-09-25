import { PDFDocument, PDFName, PDFNumber, StandardFonts, degrees, type PDFPage } from '@cantoo/pdf-lib'
import { describe, expect, it } from 'vitest'
import type { Rect } from './geometry'
import { redactPages } from './redact'
import { savePdf, extractPages } from './save'
import { extractPageTexts, extractPlacedText, fileContainsText, shownPages } from './test-support'

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

describe('saved pages as the editor shows them', () => {
  /** A 400 × 600 page with SECRET at (40, 500) and KEEP at (40, 300) in user space, in 20pt Courier (12pt per glyph). */
  async function secretPage(change: (doc: PDFDocument, page: PDFPage) => void): Promise<Uint8Array> {
    const doc = await PDFDocument.create()
    const font = await doc.embedFont(StandardFonts.Courier)
    const page = doc.addPage([400, 600])
    page.drawText('SECRET', { x: 40, y: 500, size: 20, font })
    page.drawText('KEEP', { x: 40, y: 300, size: 20, font })
    change(doc, page)
    return doc.save()
  }

  /** What Save does with a redaction box drawn on page 1 as pdf.js showed it: rebuild the pages, then redact. */
  async function saveRedacted(source: Uint8Array, box: Rect): Promise<string> {
    const [shown] = await shownPages(source)
    if (shown === undefined) throw new Error('The source has no page')
    const saved = await savePdf({ sources: [source], pages: [{ source: 0, index: 0, rotation: 0, edits: [] }] })
    const output = await redactPages(saved, [{ index: 0, shown, boxes: [box] }], { removeHiddenInformation: false })
    return (await extractPageTexts(output))[0] ?? ''
  }

  it('keeps the rotation, page box and resources a page inherits from the page tree', async () => {
    const source = await secretPage((doc, page) => {
      const { context } = doc
      const parent = context.obj({ Type: 'Pages', Parent: doc.catalog.get(PDFName.of('Pages')), Kids: [page.ref], Count: 1, Rotate: 90 })
      for (const key of ['MediaBox', 'Resources'].map(name => PDFName.of(name))) {
        const value = page.node.get(key)
        if (value !== undefined) parent.set(key, value)
        page.node.delete(key)
      }
      const parentRef = context.register(parent)
      page.node.set(PDFName.of('Parent'), parentRef)
      doc.catalog.Pages().set(PDFName.of('Kids'), context.obj([parentRef]))
    })
    expect((await shownPages(source))[0]?.rotation).toBe(90)

    // Turned a quarter clockwise, user (x, y) shows at (y, x): SECRET runs down from (500, 40).
    const text = await saveRedacted(source, { x: 490, y: 30, width: 40, height: 90 })
    expect(text).toContain('KEEP')
    expect(text).not.toContain('SECRET')
  })

  it('scales by the page\'s UserUnit, as pdf.js does', async () => {
    const source = await secretPage((_, page) => page.node.set(PDFName.of('UserUnit'), PDFNumber.of(2)))

    // Shown twice as large: SECRET's baseline starts at (80, 200).
    const text = await saveRedacted(source, { x: 70, y: 160, width: 170, height: 60 })
    expect(text).toContain('KEEP')
    expect(text).not.toContain('SECRET')
  })

  it('treats a rotation that is not a quarter turn as none, as pdf.js does', async () => {
    const source = await secretPage((_, page) => page.node.set(PDFName.of('Rotate'), PDFNumber.of(45)))
    expect((await shownPages(source))[0]?.rotation).toBe(0)

    const text = await saveRedacted(source, { x: 30, y: 75, width: 90, height: 35 })
    expect(text).toContain('KEEP')
    expect(text).not.toContain('SECRET')
  })
})

describe('extractPages', () => {
  it('keeps only the requested pages, in the requested order', async () => {
    const source = await documentWith(['ALPHA', 'BRAVO', 'CHARLIE'])
    const output = await extractPages(source, [2, 0])
    expect(await extractPageTexts(output)).toEqual(['CHARLIE', 'ALPHA'])
  })
})
