import { strFromU8, unzipSync } from 'fflate'
import {
  parse,
  type Element as SourceElement,
  type Fill,
  type Shape as SourceShape,
  type Table as SourceTable,
  type TableCell as SourceCell,
  type Text as SourceText,
} from 'pptxtojson'
import type { FindingSeverity, ImportFindingInput } from '@shared/fidelity'
import { assertZipFitsInMemory } from '../../services/zip'
import {
  defaultInset,
  lineThrough,
  mergeRuns,
  newId,
  paragraphOf,
  parseColor,
  rotatePoint,
  textRun,
  type BasicShape,
  type Border,
  type HorizontalAlign,
  type Inset,
  type ListStyle,
  type Paragraph,
  type Point,
  type Presentation,
  type RunStyle,
  type ShapeElement,
  type Slide,
  type SlideElement,
  type TableCell,
  type TableElement,
  type TextRun,
  type VerticalAlign,
} from './model'
import { defaultTheme, themes } from './themes'

/*
 * pptxtojson turns each slide into positioned elements with text as HTML, in points. This maps that
 * onto the model and reports, conservatively, everything the model cannot hold. A few things
 * pptxtojson leaves out (indent levels, bullets a placeholder inherits from its layout, highlight,
 * animations, hidden slides) are read from the slide XML directly.
 */

export interface ImportedPresentation {
  readonly presentation: Presentation
  readonly findings: readonly ImportFindingInput[]
}

const keptInOriginal = 'The original file keeps them.'

/**
 * Every finding this importer can report, worded once: what it is, and exactly what happens to it
 * here and when saving.
 */
const catalogue = {
  animations: ['Animations', 'dropped', `Slides play without them, and saving removes them. ${keptInOriginal}`],
  transitions: ['Slide transitions', 'dropped', `Slides change without them, and saving removes them. ${keptInOriginal}`],
  charts: ['Charts', 'dropped', `Not shown, and removed when saved. ${keptInOriginal}`],
  smartArt: ['SmartArt', 'dropped', 'Its shapes and text are kept as separate shapes; saving removes the SmartArt diagram itself.'],
  media: ['Video and audio', 'dropped', `Not shown, and removed when saved. ${keptInOriginal}`],
  equations: ['Equations', 'dropped', 'Shown and saved as pictures of the equations, which can no longer be edited.'],
  links: ['Links', 'dropped', 'The text, shape or picture is kept; saving removes its link.'],
  comments: ['Comments', 'dropped', `Not shown, and removed when saved. ${keptInOriginal}`],
  sections: ['Slide sections', 'dropped', 'Slides keep their order; saving removes the section names.'],
  embeddedFonts: ['Embedded fonts', 'dropped', 'Text uses fonts installed on this Mac, and saving removes the embedded copies.'],
  pictureFormats: ['Pictures in formats Slides can\'t show', 'dropped', `Not shown, and removed when saved. ${keptInOriginal}`],
  pictureFills: ['Picture fills in shapes', 'dropped', 'The shape is kept without its picture.'],
  masters: ['Slide master and layout designs', 'degraded', 'Their shapes and pictures are copied onto each slide, where they can be selected and moved.'],
  placeholders: ['Placeholders', 'degraded', 'Kept as ordinary text boxes in the same place, with the same text style.'],
  theme: ['Theme colours and fonts', 'degraded', 'Kept as fixed colours and fonts, so changing the theme in PowerPoint later does not update them.'],
  groups: ['Grouped shapes', 'degraded', 'Ungrouped into separate shapes in the same places.'],
  gradients: ['Gradient fills', 'degraded', 'Shown and saved in one colour, the gradient\'s first.'],
  patterns: ['Pattern fills', 'degraded', 'Shown and saved in one colour.'],
  gradientText: ['Gradient text', 'degraded', 'Shown and saved in one colour.'],
  shadows: ['Shadows', 'degraded', 'Shown and saved without them.'],
  textEffects: ['Strikethrough, superscript, character spacing and text shadows', 'degraded', 'Shown and saved as plain text.'],
  paragraphSpacing: ['Space before and after paragraphs', 'degraded', 'Paragraphs are shown and saved without the extra space; line spacing is kept.'],
  exactLineSpacing: ['Line spacing set in points', 'degraded', 'Converted to the nearest multiple of single spacing.'],
  listStyles: ['Custom bullets and numbering styles', 'degraded', 'Shown and saved as round bullets or 1, 2, 3.'],
  highlight: ['Text highlight in unusual places', 'degraded', 'Highlight is kept where Zendo can match it to the text, and left out elsewhere.'],
  shrinkText: ['Shrink text on overflow', 'degraded', 'Text keeps the size it was shown at, and no longer shrinks as you type.'],
  verticalText: ['Vertical text', 'degraded', 'Shown and saved horizontally.'],
  dashedLines: ['Dashed and dotted outlines', 'degraded', 'Shown and saved as solid lines.'],
  flippedShapes: ['Flipped shapes and pictures', 'degraded', 'Shown and saved unflipped.'],
  adjustedShapes: ['Adjusted shape proportions', 'degraded', 'Shown as in the file, and saved with PowerPoint\'s default proportions.'],
  connectors: ['Bent and curved connectors', 'degraded', 'Shown and saved as straight lines.'],
  doubleArrows: ['Arrowheads at both ends of a line', 'degraded', 'Kept at one end only.'],
  unknownShapes: ['Unrecognised shapes', 'degraded', 'Shown and saved as rectangles.'],
  croppedPictures: ['Cropped pictures', 'degraded', 'Shown and saved uncropped.'],
  shapedPictures: ['Pictures cut to a shape', 'degraded', 'Shown and saved as rectangles.'],
  pictureAdjustments: ['Picture brightness, contrast and colour changes', 'degraded', 'Shown and saved without them.'],
  pictureOutlines: ['Picture outlines', 'degraded', 'Shown and saved without them.'],
  hiddenSlides: ['Hidden slides', 'degraded', 'Shown and played like the other slides.'],
  fields: ['Slide numbers and dates', 'degraded', 'Kept as the text they showed when opened; they no longer update.'],
  lineBreaks: ['Line breaks inside paragraphs', 'degraded', 'Shown and saved as spaces.'],
  mergedCells: ['Merged table cells', 'degraded', 'Shown and saved as separate cells.'],
  tableText: ['Table text formatting', 'degraded', 'Each table uses one font and size; bold, colour and alignment are kept per cell.'],
  tableRuns: ['Mixed formatting inside a table cell', 'degraded', 'Each cell is shown and saved in one style, its first words\' bold and colour; highlight is left out.'],
  tableLists: ['Bullets and numbering in table cells', 'degraded', 'Shown and saved as plain lines.'],
  tableBorders: ['Table borders that differ between cells', 'degraded', 'Every cell edge is shown and saved with the same border.'],
  noWrap: ['Text set not to wrap', 'degraded', 'Shown and saved wrapping inside its box.'],
} satisfies Record<string, readonly [string, FindingSeverity, string]>

type FindingKey = keyof typeof catalogue

/** One finding per construct, listing the slides it appears on. */
class Findings {
  private readonly found = new Map<FindingKey, Set<number>>()

  add(key: FindingKey, slide?: number): void {
    const slides = this.found.get(key) ?? new Set<number>()
    if (slide !== undefined) slides.add(slide)
    this.found.set(key, slides)
  }

  list(): ImportFindingInput[] {
    return [...this.found].map(([key, slideSet]) => {
      const [construct, severity, suggestedAlternative] = catalogue[key]
      const slides = [...slideSet].sort((a, b) => a - b)
      const location = slides.length === 0 ? undefined : `${slides.length === 1 ? 'Slide' : 'Slides'} ${slides.join(', ')}`
      return { construct, severity, suggestedAlternative, ...(location === undefined ? {} : { location }) }
    })
  }
}

/* ─── What the slide XML says that pptxtojson does not ─── */

/** A stretch of a paragraph's text and its highlight colour, if any. */
interface Stretch {
  readonly text: string
  readonly highlight?: string
}

interface ParagraphInfo {
  readonly level: number
  /** Undefined when neither the paragraph nor its placeholder decides. */
  readonly list?: ListStyle
  readonly customList: boolean
  readonly stretches: readonly Stretch[]
  /** An empty paragraph's look, from its `a:endParaRPr`: pptxtojson gives empty lines a default 18 pt. */
  readonly endStyle: Partial<RunStyle>
}

/** What a table's XML says about its cells. */
interface TableXml {
  /** Each empty cell's look, from its paragraph's `a:endParaRPr`, by row and column. */
  readonly ends: readonly (readonly Partial<RunStyle>[])[]
  /** Some cell has bullets, numbering or highlighted text, none of which a cell here holds. */
  readonly lists: boolean
  readonly highlight: boolean
}

interface SlideXml {
  /** Paragraph levels, lists and highlight of each shape, by its `cNvPr` id. */
  readonly paragraphs: ReadonlyMap<string, readonly ParagraphInfo[]>
  /** Tables by their graphic frame's `cNvPr` id. */
  readonly tables: ReadonlyMap<string, TableXml>
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

function listOf(properties: Element | undefined, bulletedPlaceholder: boolean): { list?: ListStyle; customList: boolean } {
  const character = properties === undefined ? undefined : childNamed(properties, 'a:buChar')
  const numbering = properties === undefined ? undefined : childNamed(properties, 'a:buAutoNum')
  const picture = properties === undefined ? undefined : childNamed(properties, 'a:buBlip')
  const customList = (character !== undefined && !['•', '·', '▪', '–'].includes(character.getAttribute('char') ?? '•'))
    || (numbering !== undefined && (numbering.getAttribute('type') ?? 'arabicPeriod') !== 'arabicPeriod')
    || picture !== undefined
  if (numbering !== undefined) return { list: 'number', customList }
  if (character !== undefined || picture !== undefined) return { list: 'bullet', customList }
  if (properties !== undefined && childNamed(properties, 'a:buNone') !== undefined) return { list: 'none', customList }
  return { list: bulletedPlaceholder ? 'bullet' : undefined, customList }
}

/** The text of a paragraph's runs, fields and breaks in order, as pptxtojson writes them into its HTML. */
function stretchesOf(paragraph: Element): Stretch[] {
  return Array.from(paragraph.children).flatMap((child): Stretch[] => {
    if (child.tagName === 'a:br') return [{ text: ' ' }]
    if (child.tagName !== 'a:r' && child.tagName !== 'a:fld') return []
    const text = childNamed(child, 'a:t')?.textContent ?? ''
    const properties = childNamed(child, 'a:rPr')
    const highlight = properties === undefined ? undefined : childNamed(properties, 'a:highlight')
    const colour = highlight?.getElementsByTagName('a:srgbClr')[0]?.getAttribute('val')
    return [{ text, highlight: highlight === undefined ? undefined : parseColor(`#${colour ?? 'ffff00'}`) }]
  })
}

/** The run look an empty paragraph keeps in its `a:endParaRPr`: only what the file states. */
function endStyleOf(paragraph: Element): Partial<RunStyle> {
  const properties = childNamed(paragraph, 'a:endParaRPr')
  const hasRuns = Array.from(paragraph.children).some(child => child.tagName === 'a:r' || child.tagName === 'a:fld')
  if (properties === undefined || hasRuns) return {}
  const flag = (name: string): boolean | undefined => {
    const value = properties.getAttribute(name)
    return value === null ? undefined : value === '1' || value === 'true'
  }
  const size = Number(properties.getAttribute('sz'))
  const underline = properties.getAttribute('u')
  const colour = childNamed(properties, 'a:solidFill')?.getElementsByTagName('a:srgbClr')[0]
  const alpha = Number(colour?.getElementsByTagName('a:alpha')[0]?.getAttribute('val') ?? 100000)
  const alphaHex = alpha >= 100000 ? '' : Math.round((alpha / 100000) * 255).toString(16).padStart(2, '0')
  const color = colour === undefined ? undefined : parseColor(`#${colour.getAttribute('val') ?? ''}${alphaHex}`)
  const font = childNamed(properties, 'a:latin')?.getAttribute('typeface')
  const bold = flag('b')
  const italic = flag('i')
  return {
    ...(Number.isFinite(size) && size > 0 ? { size: size / 100 } : {}),
    ...(bold === undefined ? {} : { bold }),
    ...(italic === undefined ? {} : { italic }),
    ...(underline === null ? {} : { underline: underline !== 'none' }),
    ...(color === undefined ? {} : { color }),
    // A theme font (`+mn-lt`) is left to pptxtojson, which resolves it.
    ...(font === null || font === undefined || font.startsWith('+') ? {} : { font }),
  }
}

function readTableXml(frame: Element): TableXml | undefined {
  const table = frame.getElementsByTagName('a:tbl')[0]
  if (table === undefined) return undefined
  const ends = Array.from(table.children).filter(child => child.tagName === 'a:tr').map(row =>
    Array.from(row.children).filter(child => child.tagName === 'a:tc').map(cell => {
      const paragraph = cell.getElementsByTagName('a:p')[0]
      return paragraph === undefined ? {} : endStyleOf(paragraph)
    }))
  const lists = Array.from(table.getElementsByTagName('a:pPr')).some(properties =>
    childNamed(properties, 'a:buChar') !== undefined || childNamed(properties, 'a:buAutoNum') !== undefined)
  return { ends, lists, highlight: table.getElementsByTagName('a:highlight').length > 0 }
}

function readSlideXml(xml: string): SlideXml {
  const document = new DOMParser().parseFromString(xml, 'application/xml')
  const paragraphs = new Map<string, ParagraphInfo[]>()
  const tables = new Map<string, TableXml>()
  for (const frame of Array.from(document.getElementsByTagName('p:graphicFrame'))) {
    const id = frame.getElementsByTagName('p:cNvPr')[0]?.getAttribute('id')
    const table = readTableXml(frame)
    if (id !== null && id !== undefined && table !== undefined) tables.set(id, table)
  }
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
      return {
        level: Number.isFinite(level) ? Math.min(8, Math.max(0, level)) : 0,
        ...listOf(properties, bulletedPlaceholder),
        stretches: stretchesOf(paragraph),
        endStyle: endStyleOf(paragraph),
      }
    }))
  }
  const tagged = (name: string): boolean => document.getElementsByTagName(name).length > 0
  return {
    paragraphs,
    tables,
    animated: Array.from(document.getElementsByTagName('p:cTn')).some(node => node.hasAttribute('presetClass')),
    hidden: document.documentElement.getAttribute('show') === '0',
    fields: tagged('a:fld'),
    lineBreaks: tagged('a:br'),
    placeholders: tagged('p:ph'),
    // Text set in the theme's fonts; a bullet's own font (`a:buFont`) does not count.
    themed: tagged('a:schemeClr') || tagged('p:style') || /<a:(?:latin|ea|cs) typeface="\+m[jn]-/.test(xml),
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
  if (style.get('color') === 'transparent' || style.has('background')) findings.add('gradientText', slide)
  if (style.has('letter-spacing') || style.has('text-shadow') || style.has('vertical-align') || style.get('text-decoration-line') === 'line-through') {
    findings.add('textEffects', slide)
  }
  if (span.querySelector('a') !== null) findings.add('links', slide)
  const size = Number.parseFloat(style.get('font-size') ?? '18')
  // pptxtojson writes every space as a no-break space.
  return textRun((span.textContent ?? '').replace(/\u00a0/g, ' '), {
    bold: style.get('font-weight') === 'bold',
    italic: style.get('font-style') === 'italic',
    underline: (style.get('text-decoration') ?? '').includes('underline'),
    color: parseColor(style.get('color')) ?? defaultTheme.text,
    size: Math.round((Number.isFinite(size) ? size : 18) * context.fontScale * 10) / 10,
    font: (style.get('font-family') ?? '').replace(/["']/g, '').split(',')[0]?.trim() || defaultTheme.font,
  })
}

/**
 * pptxtojson writes a tab as four spaces. Where the XML has a tab and the runs have those spaces,
 * the tab comes back; if the two texts do not line up, the runs stay as they are.
 */
function withTabs(runs: readonly TextRun[], stretches: readonly Stretch[] | undefined): readonly TextRun[] {
  const xml = stretches?.map(stretch => stretch.text).join('') ?? ''
  if (!xml.includes('\t')) return runs
  let at = 0
  const restored = runs.map(run => {
    let text = ''
    for (let index = 0; index < run.text.length;) {
      if (xml[at] === '\t' && run.text.startsWith('    ', index)) {
        text += '\t'
        index += 4
      } else {
        text += run.text[index]
        index += 1
      }
      at += 1
    }
    return { ...run, text }
  })
  return at === xml.length ? restored : runs
}

/** Splits runs where the XML's highlight changes, when the two agree on the paragraph's text. */
function withHighlight(runs: readonly TextRun[], stretches: readonly Stretch[] | undefined, context: TextContext): readonly TextRun[] {
  if (stretches === undefined || stretches.every(stretch => stretch.highlight === undefined)) return runs
  if (runs.map(run => run.text).join('') !== stretches.map(stretch => stretch.text).join('')) {
    context.findings.add('highlight', context.slide)
    return runs
  }
  const pieces: TextRun[] = []
  let stretch = 0
  let left = stretches[0]?.text.length ?? 0
  for (const run of runs) {
    let text = run.text
    while (text.length > 0) {
      while (left === 0 && stretch < stretches.length - 1) left = stretches[++stretch]?.text.length ?? 0
      const take = Math.min(text.length, left > 0 ? left : text.length)
      pieces.push({ ...run, text: text.slice(0, take), highlight: stretches[stretch]?.highlight })
      text = text.slice(take)
      left -= take
    }
  }
  return mergeRuns(pieces)
}

/** CSS line height as a multiple of single spacing; pptxtojson gives a unitless multiple, or points. */
function lineSpacingOf(value: string | undefined, size: number, context: TextContext): number {
  if (value === undefined) return 1
  if (value.endsWith('pt')) {
    context.findings.add('exactLineSpacing', context.slide)
    return Math.round((Number.parseFloat(value) / (size * 1.2)) * 20) / 20 || 1
  }
  const multiple = Number.parseFloat(value)
  return Number.isFinite(multiple) && multiple > 0 ? multiple : 1
}

function readParagraphs(html: string, infos: readonly ParagraphInfo[] | undefined, context: TextContext): Paragraph[] {
  const body = new DOMParser().parseFromString(html, 'text/html').body
  return Array.from(body.querySelectorAll('p')).map((element, index) => {
    const style = styleOf(element)
    if (style.has('margin-top') || style.has('margin-bottom')) context.findings.add('paragraphSpacing', context.slide)
    const spans = Array.from(element.querySelectorAll('span')).filter(span => span.parentElement?.closest('span') === null)
    const info = infos?.[index]
    if (info?.customList === true) context.findings.add('listStyles', context.slide)
    // An empty line arrives as one no-break space.
    const empty = spans.length === 1 && spans[0]?.textContent === '\u00a0'
    const read = spans.map(span => readRun(span, context))
    // An empty line's own look is in the XML; pptxtojson only guesses it.
    const endStyle = info?.endStyle ?? {}
    const runs = read.length === 0
      ? [textRun('', { size: 18, ...endStyle })]
      : empty
        ? read.map(run => ({ ...run, ...endStyle, text: '' }))
        : withHighlight(withTabs(read, info?.stretches), info?.stretches, context)
    const htmlList: ListStyle = element.closest('ol') !== null ? 'number' : element.closest('li') !== null ? 'bullet' : 'none'
    return {
      runs,
      align: alignOf(style.get('text-align')),
      list: info?.list ?? htmlList,
      level: info?.level ?? 0,
      lineSpacing: lineSpacingOf(style.get('line-height'), runs[0]?.size ?? 18, context),
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
  readonly tables: ReadonlyMap<string, TableXml>
}

/** Layout, master and SmartArt shapes come from other parts, whose XML is not read. */
const noXml = { paragraphs: new Map<string, readonly ParagraphInfo[]>(), tables: new Map<string, TableXml>() }

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
      context.findings.add('gradients', context.slide)
      return parseColor(fill.value.colors[0]?.color)
    case 'pattern':
      context.findings.add('patterns', context.slide)
      return parseColor(fill.value.foregroundColor)
    case 'image':
      context.findings.add('pictureFills', context.slide)
      return undefined
  }
}

function borderOf(source: SourceShape | SourceText, context: ElementContext): Border | undefined {
  if (!(source.borderWidth > 0)) return undefined
  if (source.borderType !== 'solid') context.findings.add('dashedLines', context.slide)
  return { color: parseColor(source.borderColor) ?? defaultTheme.text, width: source.borderWidth }
}

function insetOf(source: SourceShape | SourceText): Inset {
  const inset = source.textInset
  return inset === undefined ? defaultInset : { top: inset.t, right: inset.r, bottom: inset.b, left: inset.l }
}

function verticalAlignOf(value: string | undefined): VerticalAlign {
  return value === 'mid' ? 'middle' : value === 'down' ? 'bottom' : 'top'
}

function textBodyOf(source: SourceShape | SourceText, context: ElementContext) {
  const fontScale = source.autoFit?.type === 'text' ? source.autoFit.fontScale ?? 1 : 1
  if (fontScale < 1) context.findings.add('shrinkText', context.slide)
  const paragraphs = source.content.trim() === ''
    ? []
    : readParagraphs(source.content, context.paragraphs.get(source.id), { findings: context.findings, slide: context.slide, fontScale })
  if (source.shadow !== undefined) context.findings.add('shadows', context.slide)
  if (source.link !== undefined) context.findings.add('links', context.slide)
  if (source.wrap === false && paragraphs.length > 0) context.findings.add('noWrap', context.slide)
  // A shape saved without text has no anchor in the file; text typed into it later starts in the middle, as in PowerPoint.
  const verticalAlign = source.type === 'shape' && paragraphs.length === 0 ? 'middle' : verticalAlignOf(source.vAlign)
  return { paragraphs, verticalAlign, inset: insetOf(source) }
}

const lineShapes: ReadonlySet<string> = new Set(['line', 'straightConnector1', 'bentConnector2', 'bentConnector3', 'bentConnector4', 'bentConnector5', 'curvedConnector2', 'curvedConnector3', 'curvedConnector4', 'curvedConnector5'])
const basicShapes: ReadonlySet<string> = new Set(['rect', 'roundRect', 'ellipse', 'triangle', 'rightArrow'])

function isBasicShape(name: string): name is BasicShape {
  return basicShapes.has(name)
}

/** Lines keep their direction in flips rather than a rotation, so their rotation is applied to their ends. */
function lineOf(source: SourceShape, context: ElementContext): ShapeElement {
  const { findings, slide } = context
  if (source.shapType !== 'line' && source.shapType !== 'straightConnector1') findings.add('connectors', slide)
  const head = source.headEnd !== undefined && source.headEnd.type !== 'none'
  const tail = source.tailEnd !== undefined && source.tailEnd.type !== 'none'
  if (head && tail) findings.add('doubleArrows', slide)
  const frame = frameOf(source, source.rotate, context)
  const centre = { x: frame.x + frame.width / 2, y: frame.y + frame.height / 2 }
  const left = frame.x
  const right = frame.x + frame.width
  const top = frame.y
  const bottom = frame.y + frame.height
  let start = rotatePoint({ x: source.isFlipH ? right : left, y: source.isFlipV ? bottom : top }, centre, frame.rotation)
  let end = rotatePoint({ x: source.isFlipH ? left : right, y: source.isFlipV ? top : bottom }, centre, frame.rotation)
  // The editor's arrow points at the end; an arrow drawn only at the start is turned around.
  if (head && !tail) [start, end] = [end, start]
  const shape: ShapeElement = {
    kind: 'shape',
    ...frame,
    geometry: { type: head || tail ? 'arrow' : 'line' },
    border: { color: parseColor(source.borderColor) ?? defaultTheme.text, width: source.borderWidth > 0 ? source.borderWidth : 1 },
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
  if (source.isFlipH || source.isFlipV) findings.add('flippedShapes', slide)
  const adjusted = source.keypoints !== undefined && Object.keys(source.keypoints).length > 0
  if (adjusted) findings.add('adjustedShapes', slide)
  const path = source.path ?? ''
  // An adjusted basic shape keeps its own outline, so it still looks as it did.
  const geometry: ShapeElement['geometry'] = isBasicShape(source.shapType) && !adjusted
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
  if (geometry.type === 'rect' && source.shapType !== 'rect') findings.add('unknownShapes', slide)
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

function isShownImage(src: string): boolean {
  return shownImageTypes.has(/^data:([^;,]*)/.exec(src)?.[1] ?? '')
}

function imageOf(src: string, source: Positioned, rotate: number, context: ElementContext): SlideElement[] {
  if (!isShownImage(src)) {
    context.findings.add('pictureFormats', context.slide)
    return []
  }
  return [{ kind: 'image', ...frameOf(source, rotate, context), src }]
}

/** A cell's text as plain lines, with the look of its first run, or of its empty paragraph's `a:endParaRPr`; a table style's colours come from pptxtojson. */
function cellOf(source: SourceCell, endStyle: Partial<RunStyle>, context: ElementContext): { readonly cell: TableCell; readonly runs: readonly TextRun[] } {
  const text: TextContext = { findings: context.findings, slide: context.slide, fontScale: 1 }
  const paragraphs = source.text.trim() === '' ? [] : readParagraphs(source.text, undefined, text)
  const runs = paragraphs.flatMap(paragraph => paragraph.runs).filter(run => run.text !== '')
  const first = runs[0]
  if (source.rowSpan !== undefined || source.colSpan !== undefined || source.hMerge !== undefined || source.vMerge !== undefined) {
    context.findings.add('mergedCells', context.slide)
  }
  // A cell holds one bold and one colour; runs that differ lose theirs.
  if (runs.some(run => run.bold !== first?.bold || run.color !== first.color)) context.findings.add('tableRuns', context.slide)
  return {
    runs,
    cell: {
      text: paragraphs.map(paragraph => paragraph.runs.map(run => run.text).join('')).join('\n'),
      bold: source.fontBold === true || (first?.bold ?? endStyle.bold ?? false),
      color: parseColor(source.fontColor) ?? first?.color ?? endStyle.color,
      fill: parseColor(source.fillColor),
      align: paragraphs[0]?.align ?? 'left',
      verticalAlign: verticalAlignOf(source.vAlign),
    },
  }
}

type SourceBorder = SourceCell['borders']['top']

function sameEdge(a: SourceBorder, b: SourceBorder): boolean {
  const width = (edge: SourceBorder): number => edge?.borderWidth ?? 0
  return width(a) === width(b) && (width(a) === 0 || parseColor(a?.borderColor) === parseColor(b?.borderColor))
}

/** The colour most cells use, so the table keeps it once and cells keep only their own. */
function commonColor(colors: readonly (string | undefined)[]): string | undefined {
  const counts = new Map<string, number>()
  for (const color of colors) if (color !== undefined) counts.set(color, (counts.get(color) ?? 0) + 1)
  return [...counts].sort((a, b) => b[1] - a[1])[0]?.[0]
}

function tableOf(source: SourceTable, context: ElementContext): TableElement {
  const xml = context.tables.get(source.id)
  if (xml?.lists === true) context.findings.add('tableLists', context.slide)
  if (xml?.highlight === true) context.findings.add('tableRuns', context.slide)
  const columns = source.colWidths.length > 0 ? source.colWidths : [source.width]
  const read = source.data.map((row, rowIndex) => row.map((cell, column) => cellOf(cell, xml?.ends[rowIndex]?.[column] ?? {}, context)))
  const runs = read.flatMap(row => row.flatMap(cell => cell.runs))
  const first = runs[0]
  if (runs.some(run => run.size !== first?.size || run.font !== first.font || run.italic || run.underline)) context.findings.add('tableText', context.slide)
  // Every cell edge takes the table's one border; an edge set differently in the file loses that.
  const edge = source.borders.top ?? source.borders.left ?? source.data[0]?.[0]?.borders.top
  const edges = source.data.flatMap(row => row.flatMap(cell => [cell.borders.top, cell.borders.bottom, cell.borders.left, cell.borders.right]))
  if (edges.some(cellEdge => cellEdge !== undefined && !sameEdge(cellEdge, edge))) context.findings.add('tableBorders', context.slide)
  const firstEnd = xml?.ends.flat().find(style => style.size !== undefined)
  const color = commonColor(read.flatMap(row => row.map(cell => cell.cell.color))) ?? defaultTheme.text
  const rows = read.map((row, index) => ({
    height: source.rowHeights[index] ?? 32,
    cells: columns.map((_, column) => {
      const cell = row[column]?.cell ?? { text: '', bold: false, align: 'left' as const, verticalAlign: 'top' as const }
      return { ...cell, color: cell.color === color ? undefined : cell.color }
    }),
  }))
  return {
    kind: 'table',
    ...frameOf(source, 0, context),
    width: columns.reduce((total, width) => total + width, 0),
    height: rows.reduce((total, row) => total + row.height, 0),
    columns,
    rows,
    font: first?.font ?? firstEnd?.font ?? defaultTheme.font,
    size: first?.size ?? firstEnd?.size ?? 18,
    color,
    border: { color: parseColor(edge?.borderColor) ?? '#9aa0a8', width: edge?.borderWidth ?? 1 },
  }
}

/** Converts one pptxtojson element; groups, SmartArt and equations become several or other elements. */
function convert(source: SourceElement, context: ElementContext): SlideElement[] {
  const { findings, slide } = context
  switch (source.type) {
    case 'text': {
      if (source.isVertical) findings.add('verticalText', slide)
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
      if (source.rect !== undefined && Object.values(source.rect).some(value => value !== undefined && value !== 0)) findings.add('croppedPictures', slide)
      if (source.geom !== undefined && source.geom !== 'rect') findings.add('shapedPictures', slide)
      if (source.filters !== undefined && Object.keys(source.filters).length > 0) findings.add('pictureAdjustments', slide)
      if (source.isFlipH || source.isFlipV) findings.add('flippedShapes', slide)
      if (source.borderWidth > 0) findings.add('pictureOutlines', slide)
      if (source.link !== undefined) findings.add('links', slide)
      return imageOf(source.base64, source, source.rotate, context)
    }
    case 'table':
      return [tableOf(source, context)]
    case 'group': {
      findings.add('groups', slide)
      const centre = { x: source.left + source.width / 2, y: source.top + source.height / 2 }
      const place: Placement = (child, rotation) =>
        context.place(rotatePoint({ x: source.left + child.x, y: source.top + child.y }, centre, source.rotate), rotation + source.rotate)
      return ordered(source.elements).flatMap(child => convert(child, { ...context, place }))
    }
    case 'diagram': {
      findings.add('smartArt', slide)
      const place: Placement = (child, rotation) => context.place({ x: source.left + child.x, y: source.top + child.y }, rotation)
      return ordered(source.elements).flatMap(child => convert(child, { ...context, ...noXml, place }))
    }
    case 'math':
      findings.add('equations', slide)
      return source.picBase64 === '' ? [] : imageOf(source.picBase64, source, 0, context)
    case 'chart':
      findings.add('charts', slide)
      return []
    case 'video':
    case 'audio':
      findings.add('media', slide)
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

/**
 * The slide files in the order the presentation shows them (its slide list), when that list names
 * exactly these files; otherwise undefined. pptxtojson reads slides in file-number order, which a
 * file edited by some tools does not keep.
 */
function presentationOrder(parts: Readonly<Record<string, Uint8Array>>, files: readonly string[]): string[] | undefined {
  const presentation = parts['ppt/presentation.xml']
  const relations = parts['ppt/_rels/presentation.xml.rels']
  if (presentation === undefined || relations === undefined) return undefined
  const targets = new Map(Array.from(strFromU8(relations).matchAll(/<Relationship\b[^>]*>/g), ([tag]) => [
    /\bId="([^"]+)"/.exec(tag)?.[1] ?? '',
    `ppt/${(/\bTarget="([^"]+)"/.exec(tag)?.[1] ?? '').replace(/^\/?ppt\//, '')}`,
  ]))
  const order = Array.from(strFromU8(presentation).matchAll(/<p:sldId\b[^>]*\br:id="([^"]+)"/g), ([, id]) => targets.get(id ?? '') ?? '')
  const same = order.length === files.length && [...order].sort().join() === [...files].sort().join()
  return same ? order : undefined
}

/** The built-in theme whose background most slides use, so a deck made with a theme keeps following it. */
function matchingTheme(slides: readonly Slide[]): string {
  const counts = new Map<string, number>()
  for (const slide of slides) if (slide.backgroundImage === undefined) counts.set(slide.background, (counts.get(slide.background) ?? 0) + 1)
  const common = [...counts].sort((a, b) => b[1] - a[1])[0]?.[0]
  return themes.find(theme => theme.background === common)?.id ?? defaultTheme.id
}

/** Reads .pptx bytes into a presentation, with an import report of what it could not keep. */
/**
 * The slide files pptxtojson reads, in its order: those the package's content types list, by file
 * number. A slide file left over in the zip, not listed there, is not a slide.
 */
function listedSlideFiles(parts: Readonly<Record<string, Uint8Array>>): string[] {
  const types = parts['[Content_Types].xml']
  const listed = types === undefined ? [] : Array.from(strFromU8(types).matchAll(/<Override\b[^>]*>/g), ([tag]) => tag)
    .filter(tag => tag.includes('presentationml.slide+xml'))
    .map(tag => (/\bPartName="\/?([^"]+)"/.exec(tag)?.[1] ?? ''))
  return listed.filter(name => name !== '').sort((a, b) => slideNumber(a) - slideNumber(b))
}

const unreadable = 'This presentation couldn\'t be read. It may be damaged, or use something Slides doesn\'t support.'

export async function readPptx(bytes: Uint8Array): Promise<ImportedPresentation> {
  assertZipFitsInMemory(bytes)
  let parts: Record<string, Uint8Array>
  try {
    parts = unzipSync(bytes, { filter: file => /\.(xml|rels)$/.test(file.name) || file.name.startsWith('ppt/comments') })
  } catch {
    throw new Error('This file is not a PowerPoint presentation, or it is damaged.')
  }
  const presentationXml = parts['ppt/presentation.xml']
  if (presentationXml === undefined) throw new Error('This file is not a PowerPoint presentation, or it is damaged.')
  let source: Awaited<ReturnType<typeof parse>>
  try {
    source = await parse(new Uint8Array(bytes).buffer, { imageMode: 'base64', videoMode: 'none', audioMode: 'none' })
  } catch {
    // pptxtojson's own errors name its internals, which say nothing to someone opening a file.
    throw new Error(unreadable)
  }
  const slideFiles = listedSlideFiles(parts)
  const findings = new Findings()

  // Pair each pptxtojson slide with its file, then put them in the presentation's own order.
  const order = presentationOrder(parts, slideFiles)
  const paired = source.slides.map((slide, index) => ({ slide, file: slideFiles[index] }))
  const inOrder = order === undefined ? paired : order.map(file => paired.find(pair => pair.file === file) ?? { slide: undefined, file })

  const slides = inOrder.flatMap(({ slide: sourceSlide, file }, index): Slide[] => {
    if (sourceSlide === undefined) return []
    const number = index + 1
    const xml = readSlideXml(file === undefined ? '<p:sld/>' : strFromU8(parts[file] ?? new Uint8Array()))
    const context: ElementContext = { findings, slide: number, place: onSlide, paragraphs: xml.paragraphs, tables: xml.tables }
    if (xml.animated) findings.add('animations', number)
    if (sourceSlide.transition !== undefined && sourceSlide.transition !== null && sourceSlide.transition.type !== 'none') findings.add('transitions', number)
    if (xml.hidden) findings.add('hiddenSlides', number)
    if (xml.fields) findings.add('fields', number)
    if (xml.lineBreaks) findings.add('lineBreaks', number)
    if (xml.placeholders) findings.add('placeholders')
    if (xml.themed) findings.add('theme')

    const fill = sourceSlide.fill
    const backgroundImage = fill.type === 'image' && isShownImage(fill.value.base64) ? fill.value.base64 : undefined
    if (fill.type === 'image' && backgroundImage === undefined) findings.add('pictureFormats', number)
    const background = fill.type === 'image' ? defaultTheme.background : fillOf(fill, context) ?? defaultTheme.background
    // Decorations from the slide's layout and master, behind everything the slide itself holds.
    const designs = ordered(sourceSlide.layoutElements).flatMap(element => convert(element, { ...context, ...noXml }))
    if (designs.length > 0) findings.add('masters', number)
    const elements = ordered(sourceSlide.elements).flatMap(element => convert(element, context))
    return [{ id: newId(), background, backgroundImage, elements: [...designs, ...elements], notes: notesText(sourceSlide.note) }]
  })

  if (source.usedFonts.length > 0) findings.add('embeddedFonts')
  if (Object.keys(parts).some(name => name.startsWith('ppt/comments'))) findings.add('comments')
  if (strFromU8(presentationXml).includes('sectionLst')) findings.add('sections')

  return {
    presentation: {
      width: source.size.width,
      height: source.size.height,
      theme: matchingTheme(slides),
      slides: slides.length > 0 ? slides : [{ id: newId(), background: defaultTheme.background, elements: [], notes: '' }],
    },
    findings: findings.list(),
  }
}
