import {
  degrees,
  LineCapStyle,
  LineJoinStyle,
  lineTo,
  moveTo,
  popGraphicsState,
  pushGraphicsState,
  rgb,
  setLineCap,
  setLineJoin,
  setLineWidth,
  setStrokingColor,
  StandardFonts,
  stroke,
  type PDFDocument,
  type PDFFont,
  type PDFPage,
  type RGB,
} from '@cantoo/pdf-lib'
import { rectToUserSpace, toUserSpace, type PageGeometry } from './geometry'
import type { PdfColor, PdfEdit } from './types'

/** Spacing between the tops of consecutive lines, as a multiple of the font size. */
const LINE_HEIGHT_RATIO = 1.2

/** Embeds Helvetica the first time a text edit actually needs it. */
export function textFontProvider(doc: PDFDocument): () => PDFFont {
  let font: PDFFont | undefined
  return () => {
    font ??= doc.embedStandardFont(StandardFonts.Helvetica)
    return font
  }
}

export function parseColor(color: PdfColor): RGB {
  const digits = color.slice(1)
  const hex = digits.length === 3 ? digits.replace(/./g, digit => digit + digit) : digits
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) throw new Error(`Unsupported PDF colour: ${color}`)
  const value = Number.parseInt(hex, 16)
  return rgb(((value >> 16) & 0xff) / 255, ((value >> 8) & 0xff) / 255, (value & 0xff) / 255)
}

function drawTextEdit(
  page: PDFPage,
  geometry: PageGeometry,
  font: PDFFont,
  edit: Extract<PdfEdit, { kind: 'text' }>,
): void {
  const lines = edit.text.replace(/\t/g, '  ').split(/\r\n|\r|\n/)
  const ascent = font.heightAtSize(edit.size, { descender: false })
  const color = parseColor(edit.color)

  lines.forEach((line, index) => {
    if (line.length === 0) return
    const baseline = edit.y + index * edit.size * LINE_HEIGHT_RATIO + ascent
    const origin = toUserSpace(geometry, { x: edit.x, y: baseline })
    page.drawText(line, {
      x: origin.x,
      y: origin.y,
      size: edit.size,
      font,
      color,
      rotate: degrees(geometry.rotation),
    })
  })
}

function drawInkEdit(
  page: PDFPage,
  geometry: PageGeometry,
  edit: Extract<PdfEdit, { kind: 'ink' }>,
): void {
  if (edit.points.length === 0) return
  const path = edit.points.map(point => toUserSpace(geometry, point))
  const [start, ...rest] = path
  const segments = rest.length > 0 ? rest : [start]

  page.pushOperators(
    pushGraphicsState(),
    setStrokingColor(parseColor(edit.color)),
    setLineWidth(edit.width),
    setLineCap(LineCapStyle.Round),
    setLineJoin(LineJoinStyle.Round),
    moveTo(start.x, start.y),
    ...segments.map(point => lineTo(point.x, point.y)),
    stroke(),
    popGraphicsState(),
  )
}

/**
 * Draws one page's edits. Everything lands in PDF user space, so content on a
 * page the viewer will rotate is turned by the same quarter turn to come out
 * upright: `degrees()` turns counter-clockwise, which cancels the clockwise
 * rotation the viewer applies.
 */
export async function applyEdits(
  doc: PDFDocument,
  page: PDFPage,
  geometry: PageGeometry,
  edits: readonly PdfEdit[],
  font: () => PDFFont,
): Promise<void> {
  for (const edit of edits) {
    if (edit.kind === 'text') {
      drawTextEdit(page, geometry, font(), edit)
    } else if (edit.kind === 'rect') {
      const box = rectToUserSpace(geometry, edit)
      page.drawRectangle({
        x: box.x,
        y: box.y,
        width: box.width,
        height: box.height,
        color: edit.fill === undefined ? undefined : parseColor(edit.fill),
        borderColor: edit.stroke === undefined ? undefined : parseColor(edit.stroke),
        borderWidth: edit.stroke === undefined ? 0 : 1,
        opacity: edit.opacity,
        borderOpacity: edit.opacity,
      })
    } else if (edit.kind === 'image') {
      const image = await doc.embedPng(edit.png)
      // pdf-lib starts a drawing from `translate -> rotate -> scale`, so the
      // anchor is the corner the image sits on: its displayed bottom-left.
      const anchor = toUserSpace(geometry, { x: edit.x, y: edit.y + edit.height })
      page.drawImage(image, {
        x: anchor.x,
        y: anchor.y,
        width: edit.width,
        height: edit.height,
        rotate: degrees(geometry.rotation),
      })
    } else {
      drawInkEdit(page, geometry, edit)
    }
  }
}
