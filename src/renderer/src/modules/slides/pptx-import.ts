import { strFromU8, unzipSync } from 'fflate'
import { parse, type Element as SourceElement, type Fill, type Shape as SourceShape, type Text as SourceText } from 'pptxtojson'
import type { FindingSeverity, ImportFindingInput } from '@shared/fidelity'
import {
  defaultBackground,
  defaultFont,
  defaultInset,
  defaultTextColor,
  lineThrough,
  newId,
  paragraphOf,
  parseColor,
  rotatePoint,
  textRun,
  type BasicShape,
  type Border,
  type HorizontalAlign,
  type Inset,
  type Paragraph,
  type Point,
  type Presentation,
  type ShapeElement,
  type Slide,
  type SlideElement,
  type TextRun,
  type VerticalAlign,
} from './model'

/*
 * pptxtojson turns each slide into positioned elements with text as HTML, in points. This maps that
 * onto the model and reports, conservatively, everything the model cannot hold. A few things
 * pptxtojson leaves out (indent levels, bullets a placeholder inherits from its layout, animations,
 * hidden slides) are read from the slide XML directly.
 */

export interface ImportedPresentation {
  readonly presentation: Presentation
  readonly findings: readonly ImportFindingInput[]
}

/** One finding per construct, listing the slides it appears on. */
class Findings {
  private readonly groups = new Map<string, { severity: FindingSeverity; slides: Set<number>; alternative?: string }>()

  add(construct: string, severity: FindingSeverity, slide?: number, alternative?: string): void {
    const group = this.groups.get(construct) ?? { severity, slides: new Set<number>(), alternative }
    if (severity === 'dropped') group.severity = 'dropped'
    if (slide !== undefined) group.slides.add(slide)
    this.groups.set(construct, group)
  }

  list(): ImportFindingInput[] {
    return [...this.groups].map(([construct, group]) => {
      const slides = [...group.slides].sort((a, b) => a - b)
      const location = slides.length === 0 ? undefined : `${slides.length === 1 ? 'Slide' : 'Slides'} ${slides.join(', ')}`
      return {
        construct,
        severity: group.severity,
        ...(location === undefined ? {} : { location }),
        ...(group.alternative === undefined ? {} : { suggestedAlternative: group.alternative }),
      }
    })
  }
}

/* ─── What the slide XML says that pptxtojson does not ─── */

interface ParagraphInfo {
  readonly level: number
  /** Undefined when neither the paragraph nor its placeholder decides. */
  readonly bullet?: boolean
}

interface SlideXml {
  /** Paragraph levels and bullets of each shape, by its `cNvPr` id. */
  readonly paragraphs: ReadonlyMap<string, readonly ParagraphInfo[]>
  readonly animated: boolean
  readonly hidden: boolean
  readonly fields: boolean
  readonly lineBreaks: boolean
  readonly placeholders: boolean
  readonly themed: boolean
}

function childNamed(parent: Element, name: string): Element | undefined {
  return Array.from(parent.children).find(child => child.tagName === name)
}

function readSlideXml(xml: string): SlideXml {
  const document = new DOMParser().parseFromString(xml, 'application/xml')
  const paragraphs = new Map<string, ParagraphInfo[]>()
  for (const shape of Array.from(document.getElementsByTagName('p:sp'))) {
    const id = shape.getElementsByTagName('p:cNvPr')[0]?.getAttribute('id')
    const body = childNamed(shape, 'p:txBody')
    if (id === null || id === undefined || body === undefined) continue
    // A content placeholder (one with no type) takes bullets from the master's body style.
    const placeholder = shape.getElementsByTagName('p:ph')[0]
    const bulletedPlaceholder = placeholder !== undefined && (placeholder.getAttribute('type') ?? 'obj') === 'obj'
    paragraphs.set(id, Array.from(body.children).filter(child => child.tagName === 'a:p').map(paragraph => {
      const properties = childNamed(paragraph, 'a:pPr')
      const level = Number(properties?.getAttribute('lvl') ?? 0)
      const explicit = properties !== undefined && (childNamed(properties, 'a:buChar') !== undefined || childNamed(properties, 'a:buAutoNum') !== undefined)
      const none = properties !== undefined && childNamed(properties, 'a:buNone') !== undefined
      const bullet = explicit ? true : none ? false : bulletedPlaceholder ? true : undefined
      return { level: Number.isFinite(level) ? Math.min(8, Math.max(0, level)) : 0, bullet }
    }))
  }
  const tagged = (name: string): boolean => document.getElementsByTagName(name).length > 0
  return {
    paragraphs,
    animated: Array.from(document.getElementsByTagName('p:cTn')).some(node => node.hasAttribute('presetClass')),
    hidden: document.documentElement.getAttribute('show') === '0',
    fields: tagged('a:fld'),
    lineBreaks: tagged('a:br'),
    placeholders: tagged('p:ph'),
    themed: tagged('a:schemeClr') || tagged('p:style') || /typeface="\+m[jn]-/.test(xml),
  }
}

/* ─── Text: pptxtojson's HTML into paragraphs and runs ─── */

/** The declarations of an inline `style` attribute, read directly so no CSS engine normalises them. */
function styleOf(element: Element): ReadonlyMap<string, string> {
  const declarations = (element.getAttribute('style') ?? '').split(';').flatMap(declaration => {
    const colon = declaration.indexOf(':')
    return colon < 0 ? [] : [[declaration.slice(0, colon).trim().toLowerCase(), declaration.slice(colon + 1).trim()] as const]
  })
  return new Map(declarations)
}

function alignOf(value: string | undefined): HorizontalAlign {
  return value === 'center' || value === 'right' || value === 'justify' ? value : 'left'
}

interface TextContext {
  readonly findings: Findings
  readonly slide: number
  /** pptxtojson reports shrink-to-fit text at full size and the scale separately. */
  readonly fontScale: number
}

function readRun(span: Element, context: TextContext): TextRun {
  const style = styleOf(span)
  const { findings, slide } = context
  const color = parseColor(style.get('color'))
  if (style.get('color') === 'transparent' || style.has('background')) findings.add('Gradient text', 'degraded', slide, 'Shown in one colour')
  if (style.has('letter-spacing') || style.has('text-shadow') || style.has('vertical-align') || style.get('text-decoration-line') === 'line-through') {
    findings.add('Text effects (strikethrough, shadows, spacing, superscript)', 'degraded', slide, 'Shown as plain text')
  }
  if (span.querySelector('a') !== null) findings.add('Links', 'dropped', slide, 'The text is kept without its link')
  const size = Number.parseFloat(style.get('font-size') ?? '18')
  return {
    // pptxtojson writes every space as a no-break space.
    text: (span.textContent ?? '').replace(/\u00a0/g, ' '),
    bold: style.get('font-weight') === 'bold',
    italic: style.get('font-style') === 'italic',
    underline: (style.get('text-decoration') ?? '').includes('underline'),
    color: color ?? defaultTextColor,
    size: Math.round((Number.isFinite(size) ? size : 18) * context.fontScale * 10) / 10,
    font: (style.get('font-family') ?? '').replace(/["']/g, '').split(',')[0]?.trim() || defaultFont,
  }
}

function readParagraphs(html: string, infos: readonly ParagraphInfo[] | undefined, context: TextContext): Paragraph[] {
  const body = new DOMParser().parseFromString(html, 'text/html').body
  return Array.from(body.querySelectorAll('p')).map((element, index) => {
    const style = styleOf(element)
    if (style.has('line-height') || style.has('margin-top') || style.has('margin-bottom')) {
      context.findings.add('Line and paragraph spacing', 'degraded', context.slide, 'Shown with single spacing')
    }
    if (element.closest('ol') !== null) context.findings.add('Numbered lists', 'degraded', context.slide, 'Shown as bullets')
    const spans = Array.from(element.querySelectorAll('span')).filter(span => span.parentElement?.closest('span') === null)
    const runs = spans.map(span => readRun(span, context))
    // An empty line arrives as one no-break space.
    const empty = runs.length === 1 && spans[0]?.textContent === '\u00a0'
    const info = infos?.[index]
    return {
      runs: runs.length === 0 ? [textRun('', { size: 18 })] : empty ? runs.map(run => ({ ...run, text: '' })) : runs,
      align: alignOf(style.get('text-align')),
      bullet: info?.bullet ?? element.closest('li') !== null,
      level: info?.level ?? 0,
    }
  })
}

function hasText(paragraphs: readonly Paragraph[]): boolean {
  return paragraphs.some(paragraph => paragraph.runs.some(run => run.text.trim() !== ''))
}

function notesText(html: string): string {
  if (html.trim() === '') return ''
  const body = new DOMParser().parseFromString(html, 'text/html').body
  const paragraphs = Array.from(body.querySelectorAll('p'))
  const lines = paragraphs.length === 0 ? [body.textContent ?? ''] : paragraphs.map(paragraph => paragraph.textContent ?? '')
  return lines.join('\n').replace(/\r\n?/g, '\n').replace(/\u00a0/g, ' ').trim()
}

/* ─── Elements ─── */

/** Where a group puts its children: maps a child's centre and rotation onto the slide. */
type Placement = (centre: Point, rotation: number) => { readonly centre: Point; readonly rotation: number }

const onSlide: Placement = (centre, rotation) => ({ centre, rotation })

interface ElementContext {
  readonly findings: Findings
  readonly slide: number
  readonly place: Placement
  readonly paragraphs: ReadonlyMap<string, readonly ParagraphInfo[]>
}

interface Positioned {
  readonly left: number
  readonly top: number
  readonly width: number
  readonly height: number
}

function frameOf(source: Positioned, rotate: number, context: ElementContext) {
  const placed = context.place({ x: source.left + source.width / 2, y: source.top + source.height / 2 }, rotate)
  return {
    id: newId(),
    x: placed.centre.x - source.width / 2,
    y: placed.centre.y - source.height / 2,
    width: source.width,
    height: source.height,
    rotation: ((placed.rotation % 360) + 360) % 360,
  }
}

/** A fill as one colour; gradients and patterns keep their main colour. */
function fillOf(fill: Fill | null | undefined, context: ElementContext): string | undefined {
  if (fill === null || fill === undefined) return undefined
  switch (fill.type) {
    case 'color':
      return parseColor(fill.value)
    case 'gradient':
      context.findings.add('Gradient fills', 'degraded', context.slide, 'Shown in one colour')
      return parseColor(fill.value.colors[0]?.color)
    case 'pattern':
      context.findings.add('Pattern fills', 'degraded', context.slide, 'Shown in one colour')
      return parseColor(fill.value.foregroundColor)
    case 'image':
      context.findings.add('Picture fills in shapes', 'degraded', context.slide, 'Left empty')
      return undefined
  }
}

function borderOf(source: SourceShape | SourceText, context: ElementContext): Border | undefined {
  if (!(source.borderWidth > 0)) return undefined
  if (source.borderType !== 'solid') context.findings.add('Dashed and dotted outlines', 'degraded', context.slide, 'Shown as solid lines')
  return { color: parseColor(source.borderColor) ?? defaultTextColor, width: source.borderWidth }
}

function insetOf(source: SourceShape | SourceText): Inset {
  const inset = source.textInset
  return inset === undefined ? defaultInset : { top: inset.t, right: inset.r, bottom: inset.b, left: inset.l }
}

function verticalAlignOf(value: string): VerticalAlign {
  return value === 'mid' ? 'middle' : value === 'down' ? 'bottom' : 'top'
}

function textBodyOf(source: SourceShape | SourceText, context: ElementContext) {
  const fontScale = source.autoFit?.type === 'text' ? source.autoFit.fontScale ?? 1 : 1
  const paragraphs = source.content.trim() === ''
    ? []
    : readParagraphs(source.content, context.paragraphs.get(source.id), { findings: context.findings, slide: context.slide, fontScale })
  if (source.shadow !== undefined) context.findings.add('Shadows', 'degraded', context.slide, 'Not shown')
  if (source.link !== undefined) context.findings.add('Links', 'dropped', context.slide, 'The shape is kept without its link')
  return { paragraphs, verticalAlign: verticalAlignOf(source.vAlign), inset: insetOf(source) }
}

const lineShapes: ReadonlySet<string> = new Set(['line', 'straightConnector1', 'bentConnector2', 'bentConnector3', 'bentConnector4', 'bentConnector5', 'curvedConnector2', 'curvedConnector3', 'curvedConnector4', 'curvedConnector5'])
const basicShapes: ReadonlySet<string> = new Set(['rect', 'roundRect', 'ellipse'])

function isBasicShape(name: string): name is BasicShape {
  return basicShapes.has(name)
}

/** Lines keep their direction in flips rather than a rotation, so their rotation is applied to their ends. */
function lineOf(source: SourceShape, context: ElementContext): ShapeElement {
  const { findings, slide } = context
  if (source.shapType !== 'line' && source.shapType !== 'straightConnector1') findings.add('Bent and curved connectors', 'degraded', slide, 'Drawn as straight lines')
  const head = source.headEnd !== undefined && source.headEnd.type !== 'none'
  const tail = source.tailEnd !== undefined && source.tailEnd.type !== 'none'
  if (head && tail) findings.add('Arrowheads at both ends', 'degraded', slide, 'Kept at one end')
  const frame = frameOf(source, source.rotate, context)
  const centre = { x: frame.x + frame.width / 2, y: frame.y + frame.height / 2 }
  const left = frame.x
  const right = frame.x + frame.width
  const top = frame.y
  const bottom = frame.y + frame.height
  let start = { x: source.isFlipH ? right : left, y: source.isFlipV ? bottom : top }
  let end = { x: source.isFlipH ? left : right, y: source.isFlipV ? top : bottom }
  start = rotatePoint(start, centre, frame.rotation)
  end = rotatePoint(end, centre, frame.rotation)
  // The editor's arrow points at the end; an arrow drawn only at the start is turned around.
  if (head && !tail) [start, end] = [end, start]
  const shape: ShapeElement = {
    kind: 'shape',
    ...frame,
    geometry: { type: head || tail ? 'arrow' : 'line' },
    border: { color: parseColor(source.borderColor) ?? defaultTextColor, width: source.borderWidth > 0 ? source.borderWidth : 1 },
    flipH: false,
    flipV: false,
    paragraphs: [],
    verticalAlign: 'middle',
    inset: defaultInset,
  }
  return lineThrough(shape, start, end)
}

function shapeOf(source: SourceShape, context: ElementContext): ShapeElement {
  if (lineShapes.has(source.shapType)) return lineOf(source, context)
  const { findings, slide } = context
  if (source.isFlipH || source.isFlipV) findings.add('Flipped shapes', 'degraded', slide, 'Shown unflipped')
  if (source.keypoints !== undefined && Object.keys(source.keypoints).length > 0) findings.add('Adjusted shape handles', 'degraded', slide, 'Saved with default proportions')
  const path = source.path ?? ''
  const geometry: ShapeElement['geometry'] = isBasicShape(source.shapType)
    ? { type: source.shapType }
    : path.trim() === ''
      ? { type: 'rect' }
      : {
          type: 'custom',
          path,
          pathWidth: source.pathViewBox?.width ?? source.width,
          pathHeight: source.pathViewBox?.height ?? source.height,
          preset: source.shapType === 'custom' ? undefined : source.shapType,
        }
  if (geometry.type === 'rect' && source.shapType !== 'rect') findings.add('Unrecognised shapes', 'degraded', slide, 'Shown as rectangles')
  return {
    kind: 'shape',
    ...frameOf(source, source.rotate, context),
    ...textBodyOf(source, context),
    geometry,
    fill: source.strokeOnly === true ? undefined : fillOf(source.fill, context),
    border: borderOf(source, context),
    flipH: false,
    flipV: false,
  }
}

const shownImageTypes: ReadonlySet<string> = new Set(['image/png', 'image/jpeg', 'image/gif'])

function imageOf(src: string, source: Positioned, rotate: number, context: ElementContext): SlideElement[] {
  const type = /^data:([^;,]*)/.exec(src)?.[1] ?? ''
  if (!shownImageTypes.has(type)) {
    context.findings.add('Pictures in formats Zendo cannot show (EMF, WMF, TIFF)', 'dropped', context.slide)
    return []
  }
  return [{ kind: 'image', ...frameOf(source, rotate, context), src }]
}

/** Converts one pptxtojson element; groups, SmartArt and equations become several or other elements. */
function convert(source: SourceElement, context: ElementContext): SlideElement[] {
  const { findings, slide } = context
  switch (source.type) {
    case 'text': {
      if (source.isVertical) findings.add('Vertical text', 'degraded', slide, 'Shown horizontally')
      const body = textBodyOf(source, context)
      return [{
        kind: 'text',
        ...frameOf(source, source.rotate, context),
        ...body,
        // An empty placeholder stays, so it can still be typed into.
        paragraphs: body.paragraphs.length === 0 ? [paragraphOf('', { size: 18 })] : body.paragraphs,
        fill: fillOf(source.fill, context),
        border: borderOf(source, context),
        prompt: hasText(body.paragraphs) ? undefined : 'Double-click to add text',
      }]
    }
    case 'shape':
      return [shapeOf(source, context)]
    case 'image': {
      if (source.rect !== undefined && Object.values(source.rect).some(value => value !== undefined && value !== 0)) findings.add('Cropped pictures', 'degraded', slide, 'Shown uncropped')
      if (source.geom !== undefined && source.geom !== 'rect') findings.add('Pictures cut to a shape', 'degraded', slide, 'Shown as rectangles')
      if (source.filters !== undefined && Object.keys(source.filters).length > 0) findings.add('Picture adjustments', 'degraded', slide, 'Shown unadjusted')
      if (source.isFlipH || source.isFlipV) findings.add('Flipped pictures', 'degraded', slide, 'Shown unflipped')
      if (source.borderWidth > 0) findings.add('Picture outlines', 'degraded', slide, 'Not shown')
      if (source.link !== undefined) findings.add('Links', 'dropped', slide, 'The picture is kept without its link')
      return imageOf(source.base64, source, source.rotate, context)
    }
    case 'group': {
      findings.add('Grouped shapes', 'degraded', slide, 'Ungrouped into separate shapes')
      const centre = { x: source.left + source.width / 2, y: source.top + source.height / 2 }
      const place: Placement = (child, rotation) =>
        context.place(rotatePoint({ x: source.left + child.x, y: source.top + child.y }, centre, source.rotate), rotation + source.rotate)
      return ordered(source.elements).flatMap(child => convert(child, { ...context, place }))
    }
    case 'diagram': {
      findings.add('SmartArt', 'dropped', slide, 'Its shapes and text are kept as separate shapes')
      const place: Placement = (child, rotation) => context.place({ x: source.left + child.x, y: source.top + child.y }, rotation)
      return ordered(source.elements).flatMap(child => convert(child, { ...context, place, paragraphs: new Map() }))
    }
    case 'math':
      findings.add('Equations', 'dropped', slide, 'Kept as pictures')
      return source.picBase64 === '' ? [] : imageOf(source.picBase64, source, 0, context)
    case 'table':
      findings.add('Tables', 'dropped', slide)
      return []
    case 'chart':
      findings.add('Charts', 'dropped', slide)
      return []
    case 'video':
    case 'audio':
      findings.add('Video and audio', 'dropped', slide)
      return []
  }
}

/** pptxtojson lists elements grouped by kind; `order` is their position in the file, back to front. */
function ordered<T extends { readonly order: number }>(elements: readonly T[]): T[] {
  return [...elements].sort((a, b) => a.order - b.order)
}

function slideNumber(name: string): number {
  return Number(/(\d+)\.xml$/.exec(name)?.[1] ?? 0)
}

/** Reads .pptx bytes into a presentation, with an import report of what it could not keep. */
export async function readPptx(bytes: Uint8Array): Promise<ImportedPresentation> {
  let parts: Record<string, Uint8Array>
  try {
    parts = unzipSync(bytes, { filter: file => /\.(xml|rels)$/.test(file.name) || file.name.startsWith('ppt/comments') })
  } catch {
    throw new Error('This file is not a PowerPoint presentation, or it is damaged.')
  }
  if (parts['ppt/presentation.xml'] === undefined) throw new Error('This file is not a PowerPoint presentation, or it is damaged.')
  const source = await parse(new Uint8Array(bytes).buffer, { imageMode: 'base64', videoMode: 'none', audioMode: 'none' })
  // pptxtojson orders slides by their file number; so does this.
  const slideFiles = Object.keys(parts).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name)).sort((a, b) => slideNumber(a) - slideNumber(b))
  const findings = new Findings()

  const slides = source.slides.map((sourceSlide, index): Slide => {
    const number = index + 1
    const file = slideFiles[index]
    const xml = readSlideXml(file === undefined ? '<p:sld/>' : strFromU8(parts[file] ?? new Uint8Array()))
    const context: ElementContext = { findings, slide: number, place: onSlide, paragraphs: xml.paragraphs }
    if (xml.animated) findings.add('Animations', 'dropped', number)
    if (sourceSlide.transition !== undefined && sourceSlide.transition !== null && sourceSlide.transition.type !== 'none') findings.add('Slide transitions', 'dropped', number)
    if (xml.hidden) findings.add('Hidden slides', 'degraded', number, 'Shown like the other slides')
    if (xml.fields) findings.add('Slide numbers and dates', 'degraded', number, 'Kept as fixed text')
    if (xml.lineBreaks) findings.add('Line breaks inside paragraphs', 'degraded', number, 'Shown as spaces')
    if (xml.placeholders) findings.add('Placeholders', 'degraded', undefined, 'Kept as text boxes with their current look')
    if (xml.themed) findings.add('Theme colours and fonts', 'degraded', undefined, 'Kept as fixed colours and fonts')

    let background = defaultBackground
    const backdrop: SlideElement[] = []
    const fill = sourceSlide.fill
    if (fill.type === 'image') {
      findings.add('Slide background pictures', 'degraded', number, 'Kept as a picture behind the slide content')
      backdrop.push(...imageOf(fill.value.base64, { left: 0, top: 0, width: source.size.width, height: source.size.height }, 0, context))
    } else {
      background = fillOf(fill, context) ?? defaultBackground
    }
    // Decorations from the slide's layout and master, behind everything the slide itself holds.
    const designs = ordered(sourceSlide.layoutElements).flatMap(element => convert(element, { ...context, paragraphs: new Map() }))
    if (designs.length > 0) findings.add('Slide master and layout designs', 'degraded', number, 'Copied onto each slide as ordinary shapes')
    const elements = ordered(sourceSlide.elements).flatMap(element => convert(element, context))
    return { id: newId(), background, elements: [...backdrop, ...designs, ...elements], notes: notesText(sourceSlide.note) }
  })

  if (source.usedFonts.length > 0) findings.add('Embedded fonts', 'dropped', undefined, 'Text uses fonts installed on this Mac')
  if (Object.keys(parts).some(name => name.startsWith('ppt/comments'))) findings.add('Comments', 'dropped')
  if (strFromU8(parts['ppt/presentation.xml']).includes('sectionLst')) findings.add('Slide sections', 'dropped', undefined, 'Slides are kept in order without section names')

  return {
    presentation: {
      width: source.size.width,
      height: source.size.height,
      slides: slides.length > 0 ? slides : [{ id: newId(), background: defaultBackground, elements: [], notes: '' }],
    },
    findings: findings.list(),
  }
}
