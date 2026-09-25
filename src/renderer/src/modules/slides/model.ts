import { defaultTheme, findTheme, type Theme } from './themes'

/**
 * A presentation as plain, serializable data. Every length is in points (1/72 inch) with the origin
 * at the slide's top-left corner: the unit PowerPoint files and pptxtojson use. Undo keeps old
 * snapshots, so nothing here is ever mutated.
 */

export type HorizontalAlign = 'left' | 'center' | 'right' | 'justify'
export type VerticalAlign = 'top' | 'middle' | 'bottom'
export type ListStyle = 'none' | 'bullet' | 'number'

export interface TextRun {
  readonly text: string
  readonly bold: boolean
  readonly italic: boolean
  readonly underline: boolean
  /** `#rrggbb`, or `#rrggbbaa` when partly transparent. */
  readonly color: string
  /** Marker colour behind the text. */
  readonly highlight?: string
  readonly size: number
  readonly font: string
}

export type RunStyle = Omit<TextRun, 'text'>

export interface Paragraph {
  /** Never empty: an empty line keeps one empty run, which carries its size and font. */
  readonly runs: readonly TextRun[]
  readonly align: HorizontalAlign
  readonly list: ListStyle
  /** 0 to 8; each level moves a list paragraph `listIndent` further right. */
  readonly level: number
  /** A multiple of single spacing: 1, 1.5, 2… */
  readonly lineSpacing: number
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

/** The shapes the editor draws itself; `rightArrow` is a block arrow, `arrow` a line with a head. */
export type BasicShape = 'rect' | 'roundRect' | 'ellipse' | 'triangle' | 'rightArrow' | 'line' | 'arrow'

/**
 * A shape imported from PowerPoint that has no tool here (a star, a callout, a freeform) keeps its
 * outline as an SVG path drawn in a `pathWidth` × `pathHeight` box and stretched to the frame,
 * plus its PowerPoint preset name to save it back as.
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

export interface TableCell {
  /** Plain text; a line break starts a new line in the cell. */
  readonly text: string
  readonly bold: boolean
  /** The table's text colour when unset. */
  readonly color?: string
  readonly fill?: string
  readonly align: HorizontalAlign
  readonly verticalAlign: VerticalAlign
}

export interface TableRow {
  /** The row's least height; a row with more text grows, as in PowerPoint. */
  readonly height: number
  readonly cells: readonly TableCell[]
}

/** A table is never rotated; its frame is the sum of its column widths and row heights. */
export interface TableElement extends Frame {
  readonly kind: 'table'
  readonly columns: readonly number[]
  readonly rows: readonly TableRow[]
  readonly font: string
  readonly size: number
  readonly color: string
  readonly border: Border
}

export type SlideElement = TextBoxElement | ShapeElement | ImageElement | TableElement

export interface Slide {
  readonly id: string
  readonly background: string
  /** A picture stretched over the whole slide, above the background colour. */
  readonly backgroundImage?: string
  /** Back to front. */
  readonly elements: readonly SlideElement[]
  readonly notes: string
}

export interface Presentation {
  readonly width: number
  readonly height: number
  /** The built-in theme new slides and objects take their colours and font from (see themes.ts). */
  readonly theme: string
  readonly slides: readonly Slide[]
}

export interface Point {
  readonly x: number
  readonly y: number
}

export interface Box {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

const titleSize = 40
const bodySize = 24
const tableTextSize = 16
/** PowerPoint's own text-box margins. */
export const defaultInset: Inset = { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 }
/** PowerPoint's own table-cell margins. */
export const cellInset: Inset = { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 }
/** Distance between indent levels and between a bullet or number and its text; what pptxgenjs writes. */
export const listIndent = 27
/** PowerPoint draws a rounded rectangle's corners at a sixth of its shorter side by default. */
export const cornerRatio = 0.16667
export const lineWidth = 2

export function newId(): string {
  return crypto.randomUUID()
}

/** A run of `text`; `style` may be a whole run, whose own text is ignored. */
export function textRun(text: string, style: Partial<TextRun> = {}): TextRun {
  return { bold: false, italic: false, underline: false, color: defaultTheme.text, size: bodySize, font: defaultTheme.font, ...style, text }
}

export function paragraphOf(text: string, style: Partial<TextRun> = {}, options: Partial<Omit<Paragraph, 'runs'>> = {}): Paragraph {
  return { runs: [textRun(text, style)], align: 'left', list: 'none', level: 0, lineSpacing: 1, ...options }
}

function textBox(box: Box, paragraph: Paragraph, verticalAlign: VerticalAlign, prompt?: string): TextBoxElement {
  return { kind: 'text', id: newId(), ...box, rotation: 0, paragraphs: [paragraph], verticalAlign, inset: defaultInset, prompt }
}

/* ─── Slide layouts ─── */

export type SlideLayout = 'title' | 'titleContent' | 'twoContent' | 'section' | 'titleOnly' | 'blank'

export const slideLayouts: ReadonlyArray<{ readonly id: SlideLayout; readonly label: string }> = [
  { id: 'title', label: 'Title' },
  { id: 'titleContent', label: 'Title and Content' },
  { id: 'twoContent', label: 'Two Content' },
  { id: 'section', label: 'Section Header' },
  { id: 'titleOnly', label: 'Title Only' },
  { id: 'blank', label: 'Blank' },
]

interface Size {
  readonly width: number
  readonly height: number
}

/** Positions are fractions of the slide, so layouts fit 16:9, 4:3 and any imported size. */
function layoutElements(layout: SlideLayout, { width, height }: Size, theme: Theme): SlideElement[] {
  const margin = width * 0.06
  const inner = width - margin * 2
  const gap = width * 0.03
  const title = { size: titleSize, color: theme.text, font: theme.font }
  const body = { size: bodySize, color: theme.text, font: theme.font }
  const subtitle = { size: bodySize, color: theme.subtle, font: theme.font }
  const heading = (y: number, h: number, verticalAlign: VerticalAlign = 'middle') =>
    textBox({ x: margin, y, width: inner, height: h }, paragraphOf('', title), verticalAlign, 'Double-click to add title')
  const content = (x: number, w: number) =>
    textBox({ x, y: height * 0.26, width: w, height: height * 0.66 }, paragraphOf('', body, { list: 'bullet' }), 'top', 'Double-click to add text')
  switch (layout) {
    case 'title':
      return [
        textBox({ x: margin, y: height * 0.26, width: inner, height: height * 0.26 }, paragraphOf('', title, { align: 'center' }), 'bottom', 'Double-click to add title'),
        textBox({ x: margin, y: height * 0.55, width: inner, height: height * 0.16 }, paragraphOf('', subtitle, { align: 'center' }), 'top', 'Double-click to add subtitle'),
      ]
    case 'titleContent':
      return [heading(height * 0.06, height * 0.16), content(margin, inner)]
    case 'twoContent':
      return [heading(height * 0.06, height * 0.16), content(margin, (inner - gap) / 2), content(margin + (inner + gap) / 2, (inner - gap) / 2)]
    case 'section':
      return [
        textBox({ x: margin, y: height * 0.3, width: inner, height: height * 0.24 }, paragraphOf('', title), 'bottom', 'Double-click to add section title'),
        textBox({ x: margin, y: height * 0.56, width: inner, height: height * 0.14 }, paragraphOf('', subtitle), 'top', 'Double-click to add text'),
      ]
    case 'titleOnly':
      return [heading(height * 0.06, height * 0.16)]
    case 'blank':
      return []
  }
}

export function createSlide(layout: SlideLayout, presentation: Pick<Presentation, 'width' | 'height' | 'theme'>): Slide {
  const theme = findTheme(presentation.theme)
  return { id: newId(), background: theme.background, elements: layoutElements(layout, presentation, theme), notes: '' }
}

/** A new presentation: 16:9 with one Title slide. */
export function newPresentation(theme: Theme = defaultTheme): Presentation {
  const size = { width: 960, height: 540, theme: theme.id }
  return { ...size, slides: [createSlide('title', size)] }
}

/* ─── New elements, placed in the middle of the slide ─── */

type Canvas = Pick<Presentation, 'width' | 'height' | 'theme'>

function centred(size: Size, width: number, height: number): Box {
  return { x: (size.width - width) / 2, y: (size.height - height) / 2, width, height }
}

export function newTextBox(canvas: Canvas): TextBoxElement {
  const theme = findTheme(canvas.theme)
  return textBox(centred(canvas, canvas.width * 0.4, 48), paragraphOf('', { color: theme.text, font: theme.font }), 'top')
}

export function newShape(shape: BasicShape, canvas: Canvas): ShapeElement {
  const theme = findTheme(canvas.theme)
  const line = shape === 'line' || shape === 'arrow'
  const width = canvas.width * 0.2
  const box = centred(canvas, width, line ? 0 : shape === 'rightArrow' ? width * 0.5 : canvas.width * 0.12)
  return {
    kind: 'shape',
    id: newId(),
    ...box,
    rotation: 0,
    geometry: { type: shape },
    fill: line ? undefined : theme.accent,
    border: line ? { color: theme.text, width: lineWidth } : undefined,
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

function emptyCell(): TableCell {
  return { text: '', bold: false, align: 'left', verticalAlign: 'top' }
}

/** A table filling most of the slide's width, with the theme's header row. */
export function newTable(rows: number, columns: number, canvas: Canvas): TableElement {
  const theme = findTheme(canvas.theme)
  const width = canvas.width * 0.8
  const rowHeight = 32
  const table: TableElement = {
    kind: 'table',
    id: newId(),
    ...centred(canvas, width, rowHeight * rows),
    rotation: 0,
    columns: Array.from({ length: columns }, () => width / columns),
    rows: Array.from({ length: rows }, () => ({ height: rowHeight, cells: Array.from({ length: columns }, emptyCell) })),
    font: theme.font,
    size: tableTextSize,
    color: theme.text,
    border: { color: '#9aa0a8', width: 1 },
  }
  return withHeaderRow(table, true, theme)
}

/* ─── Queries ─── */

export function isLine(element: SlideElement): element is ShapeElement {
  return element.kind === 'shape' && (element.geometry.type === 'line' || element.geometry.type === 'arrow')
}

/** Text boxes and every shape except lines hold text. */
export function holdsText(element: SlideElement): element is TextBoxElement | ShapeElement {
  return element.kind === 'text' || (element.kind === 'shape' && !isLine(element))
}

/** Lines keep their direction in flips and tables are never rotated. */
export function canRotate(element: SlideElement): boolean {
  return !isLine(element) && element.kind !== 'table'
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

/**
 * The number each numbered paragraph shows, as PowerPoint counts: consecutive numbered paragraphs
 * at a level count up, deeper paragraphs in between do not interrupt them, and any other paragraph
 * at that level or above starts the count again.
 */
export function listNumbers(paragraphs: readonly Pick<Paragraph, 'list' | 'level'>[]): (number | undefined)[] {
  const counts: number[] = []
  return paragraphs.map(({ list, level }) => {
    counts.length = Math.min(counts.length, level + 1)
    if (list !== 'number') {
      counts[level] = 0
      return undefined
    }
    counts[level] = (counts[level] ?? 0) + 1
    return counts[level]
  })
}

/** Every run of a text body, so toolbar buttons can show whether all of it is bold, and so on. */
export function allRuns(paragraphs: readonly Paragraph[]): readonly TextRun[] {
  return paragraphs.flatMap(paragraph => paragraph.runs)
}

function sameStyle(a: TextRun, b: TextRun): boolean {
  return a.bold === b.bold && a.italic === b.italic && a.underline === b.underline && a.color === b.color
    && a.highlight === b.highlight && a.size === b.size && a.font === b.font
}

/** Whether two text bodies hold the same paragraphs and runs, field by field. */
export function sameParagraphs(a: readonly Paragraph[], b: readonly Paragraph[]): boolean {
  return a.length === b.length && a.every((paragraph, index) => {
    const other = b[index]
    return other !== undefined && paragraph.align === other.align && paragraph.list === other.list && paragraph.level === other.level
      && paragraph.lineSpacing === other.lineSpacing && paragraph.runs.length === other.runs.length
      && paragraph.runs.every((run, position) => {
        const otherRun = other.runs[position]
        return otherRun !== undefined && run.text === otherRun.text && sameStyle(run, otherRun)
      })
  })
}

/** Joins neighbouring runs that look the same. */
export function mergeRuns(runs: readonly TextRun[]): TextRun[] {
  return runs.reduce<TextRun[]>((all, run) => {
    const last = all.at(-1)
    if (last !== undefined && sameStyle(last, run)) all[all.length - 1] = { ...last, text: last.text + run.text }
    else all.push(run)
    return all
  }, [])
}

/* ─── Changes ─── */

/**
 * Maps a list, returning the same list when every item came back unchanged. Changes that change
 * nothing then keep the presentation object, so they record no undo step and mark nothing unsaved.
 */
function mapKeeping<T>(items: readonly T[], change: (item: T) => T): readonly T[] {
  let changed = false
  const next = items.map(item => {
    const result = change(item)
    if (result !== item) changed = true
    return result
  })
  return changed ? next : items
}

export function updateSlide(presentation: Presentation, slideId: string, change: (slide: Slide) => Slide): Presentation {
  const slides = mapKeeping(presentation.slides, slide => slide.id === slideId ? change(slide) : slide)
  return slides === presentation.slides ? presentation : { ...presentation, slides }
}

/** Changes the elements `change` returns new objects for; the same slide when it returns every one unchanged. */
export function mapElements(slide: Slide, change: (element: SlideElement) => SlideElement): Slide {
  const elements = mapKeeping(slide.elements, change)
  return elements === slide.elements ? slide : { ...slide, elements }
}

export function updateElement(slide: Slide, elementId: string, change: (element: SlideElement) => SlideElement): Slide {
  return mapElements(slide, element => element.id === elementId ? change(element) : element)
}

/** Formats every run of a text body; how a selected but not edited box takes bold, size or colour. */
export function formatRuns(paragraphs: readonly Paragraph[], change: Partial<RunStyle>): Paragraph[] {
  return paragraphs.map(paragraph => ({ ...paragraph, runs: paragraph.runs.map(run => ({ ...run, ...change })) }))
}

export function translateElement<T extends SlideElement>(element: T, dx: number, dy: number): T {
  return { ...element, x: element.x + dx, y: element.y + dy }
}

export function duplicateSlide(slide: Slide): Slide {
  return { ...slide, id: newId(), elements: slide.elements.map(element => ({ ...element, id: newId() })) }
}

/** Moves elements to the front (last) or the back (first) of their slide, keeping their order among themselves. */
export function arrange(slide: Slide, ids: ReadonlySet<string>, to: 'front' | 'back'): Slide {
  const moving = slide.elements.filter(element => ids.has(element.id))
  const others = slide.elements.filter(element => !ids.has(element.id))
  const elements = to === 'front' ? [...others, ...moving] : [...moving, ...others]
  return elements.every((element, index) => element === slide.elements[index]) ? slide : { ...slide, elements }
}

/**
 * Recolours a presentation for a theme: backgrounds, and whatever text, fills and fonts still use
 * the old theme's values. Colours the user chose themselves stay.
 */
export function applyTheme(presentation: Presentation, theme: Theme): Presentation {
  const old = findTheme(presentation.theme)
  const swaps = new Map([[old.text, theme.text], [old.subtle, theme.subtle], [old.accent, theme.accent], [old.onAccent, theme.onAccent]])
  const colour = (value: string | undefined): string | undefined => value === undefined ? undefined : swaps.get(value) ?? value
  const font = (value: string): string => value === old.font ? theme.font : value
  const text = (paragraphs: readonly Paragraph[]): Paragraph[] => paragraphs.map(paragraph => ({
    ...paragraph,
    runs: paragraph.runs.map(run => ({ ...run, color: colour(run.color) ?? run.color, font: font(run.font) })),
  }))
  const element = (item: SlideElement): SlideElement => {
    switch (item.kind) {
      case 'text':
      case 'shape':
        return { ...item, paragraphs: text(item.paragraphs), fill: colour(item.fill), border: item.border && { ...item.border, color: colour(item.border.color) ?? item.border.color } }
      case 'table':
        return {
          ...item,
          font: font(item.font),
          color: colour(item.color) ?? item.color,
          rows: item.rows.map(row => ({ ...row, cells: row.cells.map(cell => ({ ...cell, color: colour(cell.color), fill: colour(cell.fill) })) })),
        }
      case 'image':
        return item
    }
  }
  return {
    ...presentation,
    theme: theme.id,
    slides: presentation.slides.map(slide => ({ ...slide, background: theme.background, elements: slide.elements.map(element) })),
  }
}

/* ─── Tables ─── */

function sum(values: readonly number[]): number {
  return values.reduce((total, value) => total + value, 0)
}

/** Keeps a table's frame equal to its columns and rows. */
function sized(table: TableElement): TableElement {
  return { ...table, width: sum(table.columns), height: sum(table.rows.map(row => row.height)) }
}

/** Stretches a table's columns and rows to a new frame size, as dragging a handle does. */
export function fitTable(table: TableElement, width: number, height: number): TableElement {
  const sx = width / Math.max(1, sum(table.columns))
  const sy = height / Math.max(1, sum(table.rows.map(row => row.height)))
  return sized({
    ...table,
    columns: table.columns.map(column => column * sx),
    rows: table.rows.map(row => ({ ...row, height: row.height * sy })),
  })
}

/** A new, empty row styled like its neighbour: the row above, or the first body row below a header. */
export function insertRow(table: TableElement, at: number): TableElement {
  const source = at === 0 ? undefined : table.rows[at === 1 && hasHeaderRow(table) ? 1 : at - 1]
  const cells = table.columns.map((_, column) => {
    const neighbour = source?.cells[column]
    return { ...emptyCell(), bold: neighbour?.bold ?? false, color: neighbour?.color, fill: neighbour?.fill, align: neighbour?.align ?? 'left' }
  })
  const row: TableRow = { height: source?.height ?? table.rows[0]?.height ?? 32, cells }
  return sized({ ...table, rows: [...table.rows.slice(0, at), row, ...table.rows.slice(at)] })
}

export function removeRow(table: TableElement, at: number): TableElement {
  if (table.rows.length <= 1) return table
  return sized({ ...table, rows: table.rows.filter((_, index) => index !== at) })
}

/** Makes room for a column by narrowing the others, so the table keeps its width. */
export function insertColumn(table: TableElement, at: number): TableElement {
  const count = table.columns.length + 1
  const narrower = (column: number): number => (column * (count - 1)) / count
  const columns = [...table.columns.slice(0, at).map(narrower), sum(table.columns) / count, ...table.columns.slice(at).map(narrower)]
  const source = Math.min(Math.max(0, at - 1), table.columns.length - 1)
  return sized({
    ...table,
    columns,
    rows: table.rows.map(row => {
      const neighbour = row.cells[source]
      const cell: TableCell = { ...emptyCell(), bold: neighbour?.bold ?? false, color: neighbour?.color, fill: neighbour?.fill }
      return { ...row, cells: [...row.cells.slice(0, at), cell, ...row.cells.slice(at)] }
    }),
  })
}

export function removeColumn(table: TableElement, at: number): TableElement {
  if (table.columns.length <= 1) return table
  const width = sum(table.columns)
  const columns = table.columns.filter((_, index) => index !== at)
  return sized({
    ...table,
    columns: columns.map(column => (column * width) / sum(columns)),
    rows: table.rows.map(row => ({ ...row, cells: row.cells.filter((_, index) => index !== at) })),
  })
}

export function updateCell(table: TableElement, row: number, column: number, change: Partial<TableCell>): TableElement {
  return {
    ...table,
    rows: table.rows.map((line, index) => index !== row ? line : {
      ...line,
      cells: line.cells.map((cell, position) => position === column ? { ...cell, ...change } : cell),
    }),
  }
}

export function setRowHeight(table: TableElement, row: number, height: number): TableElement {
  return sized({ ...table, rows: table.rows.map((line, index) => index === row ? { ...line, height } : line) })
}

/** A header row is the first row in bold on the accent colour. */
export function hasHeaderRow(table: TableElement): boolean {
  const first = table.rows[0]
  return first !== undefined && first.cells.every(cell => cell.bold && cell.fill !== undefined)
}

export function withHeaderRow(table: TableElement, on: boolean, theme: Theme): TableElement {
  const [first, ...rest] = table.rows
  if (first === undefined) return table
  const cells = first.cells.map(cell => on
    ? { ...cell, bold: true, fill: theme.accent, color: theme.onAccent }
    : { ...cell, bold: false, fill: undefined, color: undefined })
  return { ...table, rows: [{ ...first, cells }, ...rest] }
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

/** The smallest upright box around frames (rotation left out, as alignment in Keynote does). */
export function boundsOf(frames: readonly Box[]): Box | undefined {
  if (frames.length === 0) return undefined
  const left = Math.min(...frames.map(frame => frame.x))
  const top = Math.min(...frames.map(frame => frame.y))
  const right = Math.max(...frames.map(frame => frame.x + frame.width))
  const bottom = Math.max(...frames.map(frame => frame.y + frame.height))
  return { x: left, y: top, width: right - left, height: bottom - top }
}

export type AlignEdge = 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom'

/** Where an element moves to line its `edge` up with the same edge of `target`. */
export function alignTo<T extends SlideElement>(element: T, edge: AlignEdge, target: Box): T {
  switch (edge) {
    case 'left':
      return { ...element, x: target.x }
    case 'center':
      return { ...element, x: target.x + (target.width - element.width) / 2 }
    case 'right':
      return { ...element, x: target.x + target.width - element.width }
    case 'top':
      return { ...element, y: target.y }
    case 'middle':
      return { ...element, y: target.y + (target.height - element.height) / 2 }
    case 'bottom':
      return { ...element, y: target.y + target.height - element.height }
  }
}

/** Which way a resize handle faces: -1 left or top, 0 middle, 1 right or bottom. */
export interface Handle {
  readonly h: -1 | 0 | 1
  readonly v: -1 | 0 | 1
}

const minimumSize = 4

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

/**
 * The outline of the shapes drawn as polygons, in a `width` × `height` box, with PowerPoint's
 * default proportions: the triangle's apex is centred; the arrow's shaft is half the height and
 * its head half the shorter side long.
 */
export function polygonPoints(shape: 'triangle' | 'rightArrow', width: number, height: number): readonly Point[] {
  if (shape === 'triangle') return [{ x: width / 2, y: 0 }, { x: width, y: height }, { x: 0, y: height }]
  const head = Math.min(width, height) / 2
  const neck = width - head
  return [
    { x: 0, y: height / 4 }, { x: neck, y: height / 4 }, { x: neck, y: 0 }, { x: width, y: height / 2 },
    { x: neck, y: height }, { x: neck, y: (height * 3) / 4 }, { x: 0, y: (height * 3) / 4 },
  ]
}

/* ─── Colours ─── */

/** `#rrggbb` from any `#rgb`, `#rrggbb` or `#rrggbbaa`, or undefined when it is not a hex colour. */
export function opaqueHex(color: string): string | undefined {
  const match = /^#?([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.exec(color.trim())
  if (match?.[1] === undefined) return undefined
  const digits = match[1].length === 3 ? match[1].split('').map(digit => digit + digit).join('') : match[1].slice(0, 6)
  return `#${digits.toLowerCase()}`
}

/** `#rrggbb`, `#rrggbbaa` or CSS `rgb()`/`rgba()` as a model colour; undefined for anything else, and for fully transparent. */
export function parseColor(value: string | undefined): string | undefined {
  if (value === undefined) return undefined
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.exec(value.trim())
  if (hex !== null) return hex[1]?.length === 8 ? value.trim().toLowerCase() : opaqueHex(value)
  const rgb = /^rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)(?:[\s,/]+([\d.]+))?\s*\)$/i.exec(value.trim())
  if (rgb === null || Number(rgb[4] ?? 1) === 0) return undefined
  const channel = (text: string | undefined): string => Math.round(Math.min(255, Number(text ?? 0))).toString(16).padStart(2, '0')
  const alpha = rgb[4] === undefined ? 'ff' : channel(String(Number(rgb[4]) * 255))
  return `#${channel(rgb[1])}${channel(rgb[2])}${channel(rgb[3])}${alpha === 'ff' ? '' : alpha}`
}

/** Text on a fill: white on a dark one, the theme's text colour on a light one. */
export function textColorOn(fill: string | undefined, theme: Theme): string {
  const hex = fill === undefined ? undefined : opaqueHex(fill)
  if (hex === undefined) return theme.text
  const [red = 0, green = 0, blue = 0] = [1, 3, 5].map(start => Number.parseInt(hex.slice(start, start + 2), 16) / 255)
  return 0.2126 * red + 0.7152 * green + 0.0722 * blue < 0.6 ? '#ffffff' : '#222222'
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
