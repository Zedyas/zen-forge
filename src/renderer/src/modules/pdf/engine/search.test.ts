import { PDFDocument, PDFName, PDFNumber, StandardFonts, type PDFFont } from '@cantoo/pdf-lib'
import { describe, expect, it } from 'vitest'
import { buildPageText, findMatches, rangeRects, redactionBoxes, textPieces, type Advance } from '../page-text'
import { normalizeRotation, type Rect } from './geometry'
import { redactPages } from './redact'
import { savePdf } from './save'
import { searchText } from './search'
import { extractPageTexts, readTextContent, shownPages } from './test-support'
import type { PageRotation } from './types'

const line = 'Beneficiary Christopherson-Montgomery 4417229133100005 is to be paid on or by the end of a set day'
const number = '4417229133100005'

/** A page holding `content`, which can name 11pt Helvetica as /F1. */
async function pageWith(content: string, size: readonly [number, number] = [612, 792]): Promise<{ doc: PDFDocument; font: PDFFont; bytes: () => Promise<Uint8Array> }> {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const page = doc.addPage([size[0], size[1]])
  page.node.set(PDFName.of('Resources'), doc.context.obj({ Font: { F1: font.ref } }))
  page.node.set(PDFName.of('Contents'), doc.context.register(doc.context.stream(content)))
  return { doc, font, bytes: () => doc.save() }
}

/** Places characters by Helvetica's widths, as the find bar does with the font pdf.js matched. */
function measured(font: PDFFont): Advance {
  return (piece, offset) => font.widthOfTextAtSize(piece.str.slice(0, offset), 100) / font.widthOfTextAtSize(piece.str, 100)
}

const hex = (text: string): string => `<${Array.from(text, character => character.charCodeAt(0).toString(16).padStart(2, '0')).join('')}>`

/**
 * What "Redact all" and then Save do with page 1: box the matches MuPDF and pdf.js find, on the
 * page as shown turned `rotation` further, then save and redact. Resolves the page's text left.
 */
async function redactAll(bytes: Uint8Array, query: string, advance: Advance, rotation: PageRotation = 0): Promise<string> {
  const [page] = await shownPages(bytes)
  if (page === undefined) throw new Error('The page is missing')
  const shown = { ...page, rotation: normalizeRotation(page.rotation + rotation) }
  const [exact = []] = await searchText(bytes, [{ index: 0, shown }], query)
  // pdf.js's rectangles, as the find bar measures them (on the page unturned).
  let found: Rect[][] = []
  if (rotation === 0) {
    const { content, viewport } = await readTextContent(bytes)
    const text = buildPageText(textPieces(content))
    found = findMatches(text.text, query).map(range => rangeRects(text, range, viewport, advance))
  }
  const boxes = redactionBoxes(exact, found).flat()
  const saved = await savePdf({ sources: [bytes], pages: [{ source: 0, index: 0, rotation, edits: [] }] })
  const output = await redactPages(saved, [{ index: 0, shown, boxes }], { removeHiddenInformation: false })
  return (await extractPageTexts(output))[0] ?? ''
}

describe('Redact all', () => {
  it('covers a match on a line justified with word spacing (Tw), and nothing beside it', async () => {
    const { font, bytes } = await pageWith(`BT /F1 11 Tf 3 Tw 50 700 Td ${hex(line)} Tj ET`)
    const left = await redactAll(await bytes(), number, measured(font))
    expect(left).not.toMatch(/\d/)
    expect(left).toContain('Christopherson-Montgomery')
    expect(left).toContain('is to be paid')
  })

  it('covers a match on a line justified with kerning (TJ), and nothing beside it', async () => {
    const words = line.split(' ')
    const shows = words.map((word, index) => index === words.length - 1 ? hex(word) : `${hex(`${word} `)} -250`).join(' ')
    const { font, bytes } = await pageWith(`BT /F1 11 Tf 50 700 Td [${shows}] TJ ET`)
    const left = await redactAll(await bytes(), number, measured(font))
    expect(left).not.toMatch(/\d/)
    expect(left).toContain('Christopherson-Montgomery')
    expect(left).toContain('is to be paid')
  })

  it('places the boxes on a page with its own rotation, a UserUnit and a turn by the user', async () => {
    const { doc, font, bytes } = await pageWith(`BT /F1 11 Tf 40 500 Td (Account ${number} closed) Tj ET`, [400, 600])
    const [page] = doc.getPages()
    page?.node.set(PDFName.of('Rotate'), PDFNumber.of(90))
    page?.node.set(PDFName.of('UserUnit'), PDFNumber.of(1.5))
    const left = await redactAll(await bytes(), number, measured(font), 90)
    expect(left).not.toMatch(/\d/)
    expect(left).toContain('Account')
    expect(left).toContain('closed')
  })
})
