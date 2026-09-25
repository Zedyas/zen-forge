import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import PptxGenJS from 'pptxgenjs'
import {
  defaultFont,
  lineEnds,
  opaqueHex,
  scalePath,
  transparency,
  type Border,
  type Paragraph,
  type Presentation,
  type ShapeElement,
  type SlideElement,
  type TextBody,
} from './model'

/** pptxgenjs takes inches; the model is in points. */
function inches(points: number): number {
  return points / 72
}

/** pptxgenjs colours are `RRGGBB` with a separate transparency. */
function fillOf(color: string | undefined): PptxGenJS.ShapeFillProps {
  const hex = color === undefined ? undefined : opaqueHex(color)
  if (color === undefined || hex === undefined) return { type: 'none' }
  return { type: 'solid', color: hex.slice(1), transparency: transparency(color) }
}

function lineOf(border: Border | undefined, arrow: boolean): PptxGenJS.ShapeLineProps {
  const hex = border === undefined ? undefined : opaqueHex(border.color)
  if (border === undefined || hex === undefined || border.width <= 0) return { type: 'none' }
  return { type: 'solid', color: hex.slice(1), transparency: transparency(border.color), width: border.width, endArrowType: arrow ? 'triangle' : undefined }
}

/**
 * One text object per run. pptxgenjs starts a paragraph after a run with `breakLine`, and reads a
 * paragraph's bullet and level from its first run.
 */
function textRuns(paragraphs: readonly Paragraph[]): PptxGenJS.TextProps[] {
  return paragraphs.flatMap((paragraph, index) => paragraph.runs.map((run, position): PptxGenJS.TextProps => ({
    text: run.text,
    options: {
      bold: run.bold,
      italic: run.italic,
      underline: run.underline ? { style: 'sng' } : undefined,
      color: opaqueHex(run.color)?.slice(1),
      transparency: transparency(run.color),
      fontSize: run.size,
      fontFace: run.font,
      align: paragraph.align,
      ...(position === 0 ? { bullet: paragraph.bullet, indentLevel: paragraph.level } : {}),
      breakLine: position === paragraph.runs.length - 1 && index < paragraphs.length - 1,
    },
  })))
}

function frameOf(element: SlideElement): PptxGenJS.PositionProps & { rotate: number } {
  return { x: inches(element.x), y: inches(element.y), w: inches(element.width), h: inches(element.height), rotate: element.rotation }
}

function textFrameOf(body: TextBody): PptxGenJS.TextPropsOptions {
  const { inset } = body
  // pptxgenjs reads the margin array as left, right, bottom, top.
  return { valign: body.verticalAlign, margin: [inset.left, inset.right, inset.bottom, inset.top], wrap: true, fit: 'none' }
}

/** Every PowerPoint preset name pptxgenjs can write, read from its runtime enum. */
const presetNames: ReadonlySet<string> = new Set(Object.values(new PptxGenJS().ShapeType))

function isPresetName(name: string | undefined): name is PptxGenJS.SHAPE_NAME {
  return name !== undefined && presetNames.has(name)
}

/**
 * pptxgenjs writes freeforms with the `custGeom` shape (its ShapeType.custGeom), which its type
 * definitions leave out of SHAPE_NAME.
 */
const freeform = 'custGeom' as PptxGenJS.SHAPE_NAME

type FreeformPoint = NonNullable<PptxGenJS.ShapeProps['points']>[number]

/** A freeform's outline as pptxgenjs points, in inches from the shape's corner. pptxtojson writes only M, L, C, Q and Z. */
function freeformPoints(shape: ShapeElement, path: string, pathWidth: number, pathHeight: number): FreeformPoint[] {
  const scaled = scalePath(path, pathWidth === 0 ? 1 : shape.width / pathWidth, pathHeight === 0 ? 1 : shape.height / pathHeight)
  const points: FreeformPoint[] = []
  for (const [, command = '', args = ''] of scaled.matchAll(/([MLCQZ])([^MLCQZ]*)/gi)) {
    const values = args.trim().split(/[\s,]+/).filter(value => value !== '').map(value => inches(Number(value)))
    const [a = 0, b = 0, c = 0, d = 0, e = 0, f = 0] = values
    switch (command.toUpperCase()) {
      case 'M':
        points.push({ x: a, y: b, moveTo: true })
        break
      case 'L':
        points.push({ x: a, y: b })
        break
      case 'C':
        points.push({ x: e, y: f, curve: { type: 'cubic', x1: a, y1: b, x2: c, y2: d } })
        break
      case 'Q':
        points.push({ x: c, y: d, curve: { type: 'quadratic', x1: a, y1: b } })
        break
      case 'Z':
        points.push({ close: true })
        break
    }
  }
  return points
}

/** The pptxgenjs name for a shape's geometry, and its outline points when it is a freeform. */
function geometryOf(shape: ShapeElement): { readonly name: PptxGenJS.SHAPE_NAME; readonly points?: FreeformPoint[] } {
  const { geometry } = shape
  switch (geometry.type) {
    case 'rect':
    case 'roundRect':
    case 'ellipse':
    case 'line':
      return { name: geometry.type }
    case 'arrow':
      return { name: 'line' }
    case 'custom':
      if (isPresetName(geometry.preset)) return { name: geometry.preset }
      return { name: freeform, points: freeformPoints(shape, geometry.path, geometry.pathWidth, geometry.pathHeight) }
  }
}

function addShape(slide: PptxGenJS.Slide, shape: ShapeElement): void {
  const arrow = shape.geometry.type === 'arrow'
  const { name, points } = geometryOf(shape)
  const options = {
    ...frameOf(shape),
    fill: fillOf(shape.fill),
    line: lineOf(shape.border, arrow),
    flipH: shape.flipH,
    flipV: shape.flipV,
    ...(points === undefined ? {} : { points }),
  }
  if (shape.geometry.type === 'line' || shape.geometry.type === 'arrow') {
    // Lines are saved unrotated: the editor keeps their direction in the flips (see lineEnds).
    const [start, end] = lineEnds(shape)
    slide.addShape(name, { ...options, rotate: 0, flipH: start.x > end.x, flipV: start.y > end.y })
    return
  }
  if (shape.paragraphs.length === 0) {
    slide.addShape(name, options)
    return
  }
  // A shape with text is a text object with a `shape`; `points` reach the same XML writer at run time.
  const textOptions: PptxGenJS.TextPropsOptions & Pick<PptxGenJS.ShapeProps, 'points'> = { ...options, ...textFrameOf(shape), shape: name }
  slide.addText(textRuns(shape.paragraphs), textOptions)
}

function addElement(slide: PptxGenJS.Slide, element: SlideElement): void {
  switch (element.kind) {
    case 'text':
      slide.addText(textRuns(element.paragraphs), {
        ...frameOf(element),
        ...textFrameOf(element),
        isTextBox: true,
        fill: fillOf(element.fill),
        line: lineOf(element.border, false),
      })
      return
    case 'shape':
      addShape(slide, element)
      return
    case 'image':
      // pptxgenjs wants the data URL without its `data:` scheme.
      slide.addImage({ ...frameOf(element), data: element.src.replace(/^data:/, '') })
  }
}

/**
 * pptxgenjs repeats a paragraph's `<a:pPr>` before each of its runs, which the file format does not
 * allow; PowerPoint and pptxtojson then misread bullets. Keeps only the first one, and recompresses
 * (pptxgenjs stores files uncompressed).
 */
function repairParagraphProperties(bytes: Uint8Array): Uint8Array {
  const files = unzipSync(bytes)
  const repaired: Record<string, Uint8Array> = {}
  for (const [name, data] of Object.entries(files)) {
    if (name.endsWith('/')) continue
    repaired[name] = /^ppt\/slides\/slide\d+\.xml$/.test(name)
      ? strToU8(strFromU8(data).replace(/(<\/a:r>(?:<a:br\/>)?)<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g, '$1'))
      : data
  }
  return zipSync(repaired)
}

/** Writes a presentation as .pptx bytes, entirely in memory. */
export async function writePptx(presentation: Presentation): Promise<Uint8Array> {
  const pptx = new PptxGenJS()
  pptx.defineLayout({ name: 'Zendo', width: inches(presentation.width), height: inches(presentation.height) })
  pptx.layout = 'Zendo'
  pptx.theme = { headFontFace: defaultFont, bodyFontFace: defaultFont }
  for (const slide of presentation.slides) {
    const output = pptx.addSlide()
    output.background = fillOf(slide.background)
    slide.elements.forEach(element => addElement(output, element))
    if (slide.notes.trim() !== '') output.addNotes(slide.notes)
  }
  const written = await pptx.write({ outputType: 'uint8array' })
  if (!(written instanceof Uint8Array)) throw new Error('The presentation could not be written.')
  return repairParagraphProperties(written)
}
