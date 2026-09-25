/**
 * A presentation as plain, serializable data. Every length is in points (1/72 inch) with the origin
 * at the slide's top-left corner: the unit PowerPoint files and pptxtojson use. Undo keeps old
 * snapshots, so nothing here is ever mutated.
 */

export type HorizontalAlign = 'left' | 'center' | 'right' | 'justify'
export type VerticalAlign = 'top' | 'middle' | 'bottom'

export interface TextRun {
  readonly text: string
  readonly bold: boolean
  readonly italic: boolean
  readonly underline: boolean
  /** `#rrggbb`, or `#rrggbbaa` when partly transparent. */
  readonly color: string
  readonly size: number
  readonly font: string
}

export type RunStyle = Omit<TextRun, 'text'>

export interface Paragraph {
  /** Never empty: an empty line keeps one empty run, which carries its size and font. */
  readonly runs: readonly TextRun[]
  readonly align: HorizontalAlign
  readonly bullet: boolean
  /** 0 to 8; each level moves a bulleted paragraph `bulletIndent` further right. */
  readonly level: number
}

export interface Border {
  readonly color: string
  readonly width: number
}

/** Space between a text frame's edge and its text. */
export interface Inset {
  readonly top: number
  readonly right: number
  readonly bottom: number
  readonly left: number
}

export interface Frame {
  readonly id: string
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  /** Degrees clockwise around the frame's centre. */
  readonly rotation: number
}

export interface TextBody {
  /** A text box always has at least one paragraph; a shape without text has none. */
  readonly paragraphs: readonly Paragraph[]
  readonly verticalAlign: VerticalAlign
  readonly inset: Inset
}

export interface TextBoxElement extends Frame, TextBody {
  readonly kind: 'text'
  readonly fill?: string
  readonly border?: Border
  /** Shown in the editor while the box is empty ("Double-click to add title"); never saved. */
  readonly prompt?: string
}

export type BasicShape = 'rect' | 'roundRect' | 'ellipse' | 'line' | 'arrow'

/**
 * The editor draws the basic shapes itself. A shape imported from PowerPoint that has no tool here
 * (a star, a callout, a freeform) keeps its outline as an SVG path drawn in a `pathWidth` ×
 * `pathHeight` box and stretched to the frame, plus its PowerPoint preset name to save it back as.
 */
export type Geometry =
  | { readonly type: BasicShape }
  | { readonly type: 'custom'; readonly path: string; readonly pathWidth: number; readonly pathHeight: number; readonly preset?: string }

export interface ShapeElement extends Frame, TextBody {
  readonly kind: 'shape'
  readonly geometry: Geometry
  readonly fill?: string
  readonly border?: Border
  /** A line runs from its frame's top-left corner to the bottom-right one; a flip mirrors that. */
  readonly flipH: boolean
  readonly flipV: boolean
}

export interface ImageElement extends Frame {
  readonly kind: 'image'
  /** A `data:` URL, so the model stays plain data and needs no file access to draw. */
  readonly src: string
}

export type SlideElement = TextBoxElement | ShapeElement | ImageElement

export interface Slide {
  readonly id: string
  readonly background: string
  /** Back to front. */
  readonly elements: readonly SlideElement[]
  readonly notes: string
}

export interface Presentation {
  readonly width: number
  readonly height: number
  readonly slides: readonly Slide[]
}

export interface Point {
  readonly x: number
  readonly y: number
}

/* ─── Default theme: white slides, dark Arial text, which every computer can show ─── */

export const defaultFont = 'Arial'
export const defaultTextColor = '#222222'
export const secondaryTextColor = '#5f6368'
export const defaultBackground = '#ffffff'
export const titleSize = 40
export const bodySize = 24
/** PowerPoint's own text-box margins. */
export const defaultInset: Inset = { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 }
/** Distance between indent levels and between a bullet and its text; what pptxgenjs writes. */
export const bulletIndent = 27
/** PowerPoint draws a rounded rectangle's corners at a sixth of its shorter side by default. */
export const cornerRatio = 0.16667
export const lineWidth = 2

export function newId(): string {
  return crypto.randomUUID()
}

export function textRun(text: string, style: Partial<RunStyle> = {}): TextRun {
  return { text, bold: false, italic: false, underline: false, color: defaultTextColor, size: bodySize, font: defaultFont, ...style }
}

export function paragraphOf(text: string, style: Partial<RunStyle> = {}, options: Partial<Omit<Paragraph, 'runs'>> = {}): Paragraph {
  return { runs: [textRun(text, style)], align: 'left', bullet: false, level: 0, ...options }
}

interface Box {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

function textBox(box: Box, paragraph: Paragraph, verticalAlign: VerticalAlign, prompt?: string): TextBoxElement {
  return { kind: 'text', id: newId(), ...box, rotation: 0, paragraphs: [paragraph], verticalAlign, inset: defaultInset, prompt }
}

/* ─── Slide layouts ─── */

export type SlideLayout = 'title' | 'titleContent' | 'section' | 'blank'

export const slideLayouts: ReadonlyArray<{ readonly id: SlideLayout; readonly label: string }> = [
  { id: 'title', label: 'Title' },
  { id: 'titleContent', label: 'Title and Content' },
  { id: 'section', label: 'Section Header' },
  { id: 'blank', label: 'Blank' },
]

interface Size {
  readonly width: number
  readonly height: number
}

/** Positions are fractions of the slide, so layouts fit 16:9, 4:3 and any imported size. */
function layoutElements(layout: SlideLayout, { width, height }: Size): SlideElement[] {
  const margin = width * 0.06
  const inner = width - margin * 2
  const title = (size = titleSize) => ({ size, color: defaultTextColor })
  const subtitle = { size: bodySize, color: secondaryTextColor }
  switch (layout) {
    case 'title':
      return [
        textBox({ x: margin, y: height * 0.26, width: inner, height: height * 0.26 }, paragraphOf('', title(), { align: 'center' }), 'bottom', 'Double-click to add title'),
        textBox({ x: margin, y: height * 0.55, width: inner, height: height * 0.16 }, paragraphOf('', subtitle, { align: 'center' }), 'top', 'Double-click to add subtitle'),
      ]
    case 'titleContent':
      return [
        textBox({ x: margin, y: height * 0.06, width: inner, height: height * 0.16 }, paragraphOf('', title()), 'middle', 'Double-click to add title'),
        textBox({ x: margin, y: height * 0.26, width: inner, height: height * 0.66 }, paragraphOf('', { size: bodySize }, { bullet: true }), 'top', 'Double-click to add text'),
      ]
    case 'section':
      return [
        textBox({ x: margin, y: height * 0.3, width: inner, height: height * 0.24 }, paragraphOf('', title()), 'bottom', 'Double-click to add section title'),
        textBox({ x: margin, y: height * 0.56, width: inner, height: height * 0.14 }, paragraphOf('', subtitle), 'top', 'Double-click to add text'),
      ]
    case 'blank':
      return []
  }
}

export function createSlide(layout: SlideLayout, size: Size): Slide {
  return { id: newId(), background: defaultBackground, elements: layoutElements(layout, size), notes: '' }
}

/** A new presentation: 16:9 with one Title slide. */
export function newPresentation(): Presentation {
  const size = { width: 960, height: 540 }
  return { ...size, slides: [createSlide('title', size)] }
}

/* ─── New elements, placed in the middle of the slide ─── */

function centred(size: Size, width: number, height: number): Box {
  return { x: (size.width - width) / 2, y: (size.height - height) / 2, width, height }
}

export function newTextBox(size: Size): TextBoxElement {
  return textBox(centred(size, size.width * 0.4, 48), paragraphOf(''), 'top')
}

export function newShape(shape: BasicShape, size: Size): ShapeElement {
  const line = shape === 'line' || shape === 'arrow'
  const box = centred(size, size.width * 0.2, line ? 0 : size.width * 0.12)
  return {
    kind: 'shape',
    id: newId(),
    ...box,
    rotation: 0,
    geometry: { type: shape },
    fill: line ? undefined : '#4a78c2',
    border: line ? { color: defaultTextColor, width: lineWidth } : undefined,
    flipH: false,
    flipV: false,
    paragraphs: [],
    verticalAlign: 'middle',
    inset: defaultInset,
  }
}

/** An image at its natural size, shrunk to fit within most of the slide. */
export function newImage(src: string, natural: Size, size: Size): ImageElement {
  const scale = Math.min(1, (size.width * 0.8) / natural.width, (size.height * 0.8) / natural.height)
  return { kind: 'image', id: newId(), ...centred(size, natural.width * scale, natural.height * scale), rotation: 0, src }
}

/* ─── Queries ─── */

export function isLine(element: SlideElement): element is ShapeElement {
  return element.kind === 'shape' && (element.geometry.type === 'line' || element.geometry.type === 'arrow')
}

/** Text boxes and every shape except lines hold text. */
export function holdsText(element: SlideElement): element is TextBoxElement | ShapeElement {
  return element.kind === 'text' || (element.kind === 'shape' && !isLine(element))
}

export function plainText(paragraphs: readonly Paragraph[]): string {
  return paragraphs.map(paragraph => paragraph.runs.map(run => run.text).join('')).join('\n')
}

export function findSlide(presentation: Presentation, slideId: string): Slide | undefined {
  return presentation.slides.find(slide => slide.id === slideId)
}

export function findElement(slide: Slide | undefined, elementId: string | undefined): SlideElement | undefined {
  return slide?.elements.find(element => element.id === elementId)
}

/** Every run of a text body, so toolbar buttons can show whether all of it is bold, and so on. */
export function allRuns(paragraphs: readonly Paragraph[]): readonly TextRun[] {
  return paragraphs.flatMap(paragraph => paragraph.runs)
}

/* ─── Changes ─── */

export function updateSlide(presentation: Presentation, slideId: string, change: (slide: Slide) => Slide): Presentation {
  return { ...presentation, slides: presentation.slides.map(slide => slide.id === slideId ? change(slide) : slide) }
}

export function updateElement(slide: Slide, elementId: string, change: (element: SlideElement) => SlideElement): Slide {
  return { ...slide, elements: slide.elements.map(element => element.id === elementId ? change(element) : element) }
}

/** Formats every run of a text body; how a selected but not edited box takes bold, size or colour. */
export function formatRuns(paragraphs: readonly Paragraph[], change: Partial<RunStyle>): Paragraph[] {
  return paragraphs.map(paragraph => ({ ...paragraph, runs: paragraph.runs.map(run => ({ ...run, ...change })) }))
}

export function translateElement(element: SlideElement, dx: number, dy: number): SlideElement {
  return { ...element, x: element.x + dx, y: element.y + dy }
}

export function duplicateElement(element: SlideElement, offset: number): SlideElement {
  return { ...element, id: newId(), x: element.x + offset, y: element.y + offset }
}

export function duplicateSlide(slide: Slide): Slide {
  return { ...slide, id: newId(), elements: slide.elements.map(element => ({ ...element, id: newId() })) }
}

/** Moves an element to the front (last) or the back (first) of its slide. */
export function arrange(slide: Slide, elementId: string, to: 'front' | 'back'): Slide {
  const element = findElement(slide, elementId)
  if (element === undefined) return slide
  const others = slide.elements.filter(candidate => candidate.id !== elementId)
  return { ...slide, elements: to === 'front' ? [...others, element] : [element, ...others] }
}

/* ─── Geometry ─── */

export function rotatePoint(point: Point, centre: Point, degrees: number): Point {
  const radians = (degrees * Math.PI) / 180
  const cos = Math.cos(radians)
  const sin = Math.sin(radians)
  const dx = point.x - centre.x
  const dy = point.y - centre.y
  return { x: centre.x + dx * cos - dy * sin, y: centre.y + dx * sin + dy * cos }
}

/** Which way a resize handle faces: -1 left or top, 0 middle, 1 right or bottom. */
export interface Handle {
  readonly h: -1 | 0 | 1
  readonly v: -1 | 0 | 1
}

export const minimumSize = 4

/**
 * Resizes a (possibly rotated) frame by dragging one handle `delta` points on the slide. The
 * handle's opposite edge or corner stays where it is on screen, as in Keynote and PowerPoint.
 */
export function resizeFrame<T extends Frame>(frame: T, handle: Handle, delta: Point, keepAspect: boolean): T {
  const local = rotatePoint(delta, { x: 0, y: 0 }, -frame.rotation)
  let width = handle.h === 0 ? frame.width : Math.max(minimumSize, frame.width + handle.h * local.x)
  let height = handle.v === 0 ? frame.height : Math.max(minimumSize, frame.height + handle.v * local.y)
  if (keepAspect && frame.width > 0 && frame.height > 0) {
    const ratio = frame.width / frame.height
    if (handle.v === 0 || (handle.h !== 0 && width / frame.width > height / frame.height)) height = width / ratio
    else width = height * ratio
  }
  const centre = { x: frame.x + frame.width / 2, y: frame.y + frame.height / 2 }
  const anchor = rotatePoint({ x: centre.x - (handle.h * frame.width) / 2, y: centre.y - (handle.v * frame.height) / 2 }, centre, frame.rotation)
  const offset = rotatePoint({ x: (handle.h * width) / 2, y: (handle.v * height) / 2 }, { x: 0, y: 0 }, frame.rotation)
  const next = { x: anchor.x + offset.x, y: anchor.y + offset.y }
  return { ...frame, x: next.x - width / 2, y: next.y - height / 2, width, height }
}

/** A line's two ends on the slide, following its flips. Lines are never rotated (see `lineThrough`). */
export function lineEnds(shape: ShapeElement): readonly [Point, Point] {
  const left = shape.x
  const right = shape.x + shape.width
  const top = shape.y
  const bottom = shape.y + shape.height
  return [
    { x: shape.flipH ? right : left, y: shape.flipV ? bottom : top },
    { x: shape.flipH ? left : right, y: shape.flipV ? top : bottom },
  ]
}

/** The same line redrawn from `start` to `end`, as a frame with flips. */
export function lineThrough(shape: ShapeElement, start: Point, end: Point): ShapeElement {
  return {
    ...shape,
    x: Math.min(start.x, end.x),
    y: Math.min(start.y, end.y),
    width: Math.abs(end.x - start.x),
    height: Math.abs(end.y - start.y),
    rotation: 0,
    flipH: start.x > end.x,
    flipV: start.y > end.y,
  }
}

/* ─── Colours ─── */

/** `#rrggbb` from any `#rgb`, `#rrggbb` or `#rrggbbaa`, or undefined when it is not a hex colour. */
export function opaqueHex(color: string): string | undefined {
  const match = /^#?([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.exec(color.trim())
  if (match?.[1] === undefined) return undefined
  const digits = match[1].length === 3 ? match[1].split('').map(digit => digit + digit).join('') : match[1].slice(0, 6)
  return `#${digits.toLowerCase()}`
}

/** `#rrggbb`, `#rrggbbaa` or CSS `rgb()`/`rgba()` as a model colour; undefined for anything else. */
export function parseColor(value: string | undefined): string | undefined {
  if (value === undefined) return undefined
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.exec(value.trim())
  if (hex !== null) return hex[1]?.length === 8 ? value.trim().toLowerCase() : opaqueHex(value)
  const rgb = /^rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)(?:[\s,/]+([\d.]+))?\s*\)$/i.exec(value.trim())
  if (rgb === null) return undefined
  const channel = (text: string | undefined): string => Math.round(Math.min(255, Number(text ?? 0))).toString(16).padStart(2, '0')
  const alpha = rgb[4] === undefined ? 'ff' : channel(String(Number(rgb[4]) * 255))
  return `#${channel(rgb[1])}${channel(rgb[2])}${channel(rgb[3])}${alpha === 'ff' ? '' : alpha}`
}

/** Text typed into an empty shape: 18pt like PowerPoint, white on a dark fill and dark on a light one. */
export function shapeTextStyle(fill: string | undefined): Partial<RunStyle> {
  const hex = fill === undefined ? undefined : opaqueHex(fill)
  if (hex === undefined) return { size: 18 }
  const [red = 0, green = 0, blue = 0] = [1, 3, 5].map(start => Number.parseInt(hex.slice(start, start + 2), 16) / 255)
  const luminance = 0.2126 * red + 0.7152 * green + 0.0722 * blue
  return { size: 18, color: luminance < 0.6 ? '#ffffff' : defaultTextColor }
}

/** 0 (opaque) to 100 (invisible), the way PowerPoint states transparency. */
export function transparency(color: string): number {
  const match = /^#[0-9a-f]{6}([0-9a-f]{2})$/i.exec(color.trim())
  return match?.[1] === undefined ? 0 : Math.round((1 - Number.parseInt(match[1], 16) / 255) * 100)
}

/**
 * Scales an SVG path's coordinates, so a shape's outline follows its frame without stretching its
 * outline stroke. Handles every command pptxtojson writes, absolute or relative.
 */
export function scalePath(path: string, sx: number, sy: number): string {
  const tokens = path.match(/[a-zA-Z]|-?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/g) ?? []
  // Per command, whether each argument is an x (x), a y (y) or left alone (n).
  const kinds: Record<string, string> = { M: 'xy', L: 'xy', T: 'xy', H: 'x', V: 'y', C: 'xyxyxy', S: 'xyxy', Q: 'xyxy', A: 'xynnnxy', Z: '' }
  let out = ''
  let pattern = ''
  let index = 0
  for (const token of tokens) {
    if (/[a-zA-Z]/.test(token)) {
      pattern = kinds[token.toUpperCase()] ?? ''
      index = 0
      out += ` ${token}`
      continue
    }
    const kind = pattern.length === 0 ? 'n' : pattern[index % pattern.length]
    const value = Number(token)
    out += ` ${kind === 'x' ? value * sx : kind === 'y' ? value * sy : value}`
    index += 1
  }
  return out.trim()
}
