/**
 * Reads .docx into a Sumi document. The package is size-checked, only the parts this reader uses
 * are unzipped (fflate), and their XML is parsed with DOMParser, then mapped to TipTap JSON:
 * - styles.xml becomes the document's theme (Normal, Title, Heading 1–3, Quote);
 * - formatting set on the text itself (direct formatting) becomes marks and paragraph attributes,
 *   as does the formatting of any other paragraph style;
 * - numbering.xml says which paragraphs are bulleted, numbered or checklist items, and at what level;
 * - the last section gives the page setup; every section's headers, footers and columns are checked.
 * Whatever is not kept is reported in the import report, with what happens to it on save. Anything
 * unexpected is reported too, so saving asks first; the save also re-reads what it writes
 * (doc-documents.ts), which catches what these checks miss.
 */

import type { JSONContent } from '@tiptap/core'
import { strFromU8, unzipSync } from 'fflate'
import type { FindingSeverity, ImportFindingInput } from '@shared/fidelity'
import { assertZipFitsInMemory } from '../../../services/zip'
import {
  contentWidthPixels,
  defaultPage,
  defaultTheme,
  singleLineHeight,
  type BlockStyle,
  type DocumentProperties,
  type DocumentTheme,
  type PageSetup,
  type StyleName,
} from '../theme'
import { checklistMarks, docxStyleNames, highlightColors, safeHref } from './docx-styles'
import { toDataUrl } from './images'

export interface DocxRead {
  readonly content: JSONContent
  readonly findings: readonly ImportFindingInput[]
}

// ─── Package ────────────────────────────────────────────────────────────

const xmlParts = new Set([
  'word/document.xml', 'word/styles.xml', 'word/numbering.xml', 'word/comments.xml', 'word/settings.xml',
  'word/theme/theme1.xml', 'word/_rels/document.xml.rels', 'docProps/core.xml',
])
const headerOrFooter = /^word\/(header|footer)\d*\.xml$/

/** Inflates only the entries `wanted` accepts; the rest of the package is never unpacked. */
function unzipParts(bytes: Uint8Array, wanted: (name: string) => boolean): Record<string, Uint8Array> {
  return unzipSync(bytes, { filter: file => wanted(file.name) })
}

// ─── XML ────────────────────────────────────────────────────────────────

function parseXml(text: string): Element | undefined {
  const document = new DOMParser().parseFromString(text, 'application/xml')
  return document.getElementsByTagName('parsererror').length > 0 ? undefined : document.documentElement
}

function kids(element: Element | undefined, name?: string): Element[] {
  return element === undefined ? [] : Array.from(element.children).filter(child => name === undefined || child.tagName === name)
}

function kid(element: Element | undefined, name: string): Element | undefined {
  return kids(element, name)[0]
}

function attr(element: Element | undefined, name: string): string | undefined {
  return element?.getAttribute(name) ?? undefined
}

/** The `w:val` of a property element such as `<w:jc w:val="center"/>`. */
function val(parent: Element | undefined, name: string): string | undefined {
  return attr(kid(parent, name), 'w:val')
}

function all(element: Element | undefined, name: string): Element[] {
  return element === undefined ? [] : Array.from(element.getElementsByTagName(name))
}

/** A toggle such as `<w:b/>` is on unless its value says otherwise; undefined when absent. */
function toggle(parent: Element | undefined, name: string): boolean | undefined {
  const element = kid(parent, name)
  if (element === undefined) return undefined
  const value = attr(element, 'w:val')
  return value === undefined || !['0', 'false', 'off', 'none'].includes(value)
}

function number(value: string | undefined): number | undefined {
  const parsed = value === undefined ? Number.NaN : Number(value)
  return Number.isFinite(parsed) ? parsed : undefined
}

function round(value: number): number {
  return Math.round(value * 100) / 100
}

/**
 * Children of `name` inside a parent, looking through content controls and custom XML, which can
 * wrap table rows and cells as they wrap paragraphs.
 */
function wrapped(parent: Element, name: string): Element[] {
  return kids(parent).flatMap(child => {
    if (child.tagName === name) return [child]
    if (child.tagName === 'w:sdt') return wrapped(kid(child, 'w:sdtContent') ?? child, name)
    if (child.tagName === 'w:customXml') return wrapped(child, name)
    return []
  })
}

// ─── Findings ───────────────────────────────────────────────────────────

class Findings {
  private readonly entries = new Map<string, { severity: FindingSeverity; count: number; unit?: string; alternative?: string; names: Set<string> }>()

  add(construct: string, severity: FindingSeverity, options: { readonly unit?: string; readonly alternative?: string; readonly name?: string; readonly count?: number } = {}): void {
    const entry = this.entries.get(construct) ?? { severity, count: 0, unit: options.unit, alternative: options.alternative, names: new Set<string>() }
    entry.count += options.count ?? 1
    if (options.name !== undefined) entry.names.add(options.name)
    this.entries.set(construct, entry)
  }

  list(): ImportFindingInput[] {
    return [...this.entries].map(([construct, entry]) => {
      const location = entry.names.size > 0
        ? [...entry.names].join(', ')
        : entry.unit === undefined ? undefined : `${entry.count} ${entry.count === 1 ? entry.unit : `${entry.unit}s`}`
      return {
        construct,
        severity: entry.severity,
        ...(location === undefined ? {} : { location }),
        ...(entry.alternative === undefined ? {} : { suggestedAlternative: entry.alternative }),
      }
    })
  }
}

const notKept = 'Not shown, and not kept when saved.'

// ─── Formatting properties ──────────────────────────────────────────────

/** Only the properties an element sets are present, so spreading one set over another overrides just those. */
interface RunProps {
  bold?: boolean
  italic?: boolean
  underline?: string
  strike?: boolean
  color?: string
  background?: string
  size?: number
  font?: string
  vertAlign?: string
  hidden?: boolean
  effects?: boolean
}

interface ParagraphProps {
  align?: string
  spaceBefore?: number
  spaceAfter?: number
  lineHeight?: number
  /** Exact or minimum line spacing, in points; converted once the font size is known. */
  lineExact?: number
  indentLeft?: number
  indentFirstLine?: number
  pageBreakBefore?: boolean
  numId?: string
  level?: number
  outlineLevel?: number
  decorated?: boolean
  tabStops?: boolean
  rightToLeft?: boolean
}

function assign<T extends object, K extends keyof T>(target: T, key: K, value: T[K] | undefined): void {
  if (value !== undefined) target[key] = value
}

function readRunProps(rPr: Element | undefined, themeFonts: ThemeFonts): RunProps {
  if (rPr === undefined) return {}
  const fonts = kid(rPr, 'w:rFonts')
  const themeFont = attr(fonts, 'w:asciiTheme') ?? attr(fonts, 'w:hAnsiTheme')
  const color = val(rPr, 'w:color')
  const highlight = val(rPr, 'w:highlight')
  const shading = attr(kid(rPr, 'w:shd'), 'w:fill')
  const size = number(val(rPr, 'w:sz'))
  const effects = ['w:caps', 'w:smallCaps', 'w:shadow', 'w:outline', 'w:emboss', 'w:imprint', 'w:rtl'].some(name => toggle(rPr, name) === true)
    || kid(rPr, 'w:spacing') !== undefined || kid(rPr, 'w:w') !== undefined || kid(rPr, 'w:position') !== undefined
  const props: RunProps = {}
  assign(props, 'bold', toggle(rPr, 'w:b'))
  assign(props, 'italic', toggle(rPr, 'w:i'))
  assign(props, 'underline', val(rPr, 'w:u') ?? (kid(rPr, 'w:u') === undefined ? undefined : 'single'))
  assign(props, 'strike', toggle(rPr, 'w:strike') ?? toggle(rPr, 'w:dstrike'))
  assign(props, 'color', color === undefined || color === 'auto' ? undefined : `#${color.toLowerCase()}`)
  assign(props, 'background', highlight !== undefined && highlight !== 'none'
    ? highlightColors[highlight]
    : shading === undefined || shading === 'auto' ? undefined : `#${shading.toLowerCase()}`)
  assign(props, 'size', size === undefined ? undefined : size / 2)
  assign(props, 'font', attr(fonts, 'w:ascii') ?? attr(fonts, 'w:hAnsi') ?? (themeFont === undefined ? undefined : themeFont.startsWith('major') ? themeFonts.major : themeFonts.minor))
  assign(props, 'vertAlign', val(rPr, 'w:vertAlign'))
  assign(props, 'hidden', toggle(rPr, 'w:vanish'))
  if (effects) props.effects = true
  return props
}

const alignments: Readonly<Record<string, string>> = {
  left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify',
}

function readParagraphProps(pPr: Element | undefined): ParagraphProps {
  if (pPr === undefined) return {}
  const spacing = kid(pPr, 'w:spacing')
  const indent = kid(pPr, 'w:ind')
  const line = number(attr(spacing, 'w:line'))
  const rule = attr(spacing, 'w:lineRule') ?? 'auto'
  const hanging = number(attr(indent, 'w:hanging'))
  const firstLine = number(attr(indent, 'w:firstLine'))
  const numPr = kid(pPr, 'w:numPr')
  const align = val(pPr, 'w:jc')
  const before = number(attr(spacing, 'w:before'))
  const after = number(attr(spacing, 'w:after'))
  const left = number(attr(indent, 'w:left') ?? attr(indent, 'w:start'))
  const props: ParagraphProps = {}
  assign(props, 'align', align === undefined ? undefined : alignments[align])
  assign(props, 'spaceBefore', before === undefined ? undefined : before / 20)
  assign(props, 'spaceAfter', after === undefined ? undefined : after / 20)
  assign(props, 'lineHeight', line !== undefined && rule === 'auto' ? round(line / 240) : undefined)
  assign(props, 'lineExact', line !== undefined && rule !== 'auto' ? line / 20 : undefined)
  assign(props, 'indentLeft', left === undefined ? undefined : left / 20)
  assign(props, 'indentFirstLine', hanging !== undefined ? -hanging / 20 : firstLine === undefined ? undefined : firstLine / 20)
  assign(props, 'pageBreakBefore', toggle(pPr, 'w:pageBreakBefore'))
  assign(props, 'numId', val(numPr, 'w:numId'))
  assign(props, 'level', number(val(numPr, 'w:ilvl')))
  assign(props, 'outlineLevel', number(val(pPr, 'w:outlineLvl')))
  assign(props, 'rightToLeft', toggle(pPr, 'w:bidi'))
  if (kid(pPr, 'w:pBdr') !== undefined || (attr(kid(pPr, 'w:shd'), 'w:fill') ?? 'auto') !== 'auto') props.decorated = true
  if (kid(pPr, 'w:tabs') !== undefined) props.tabStops = true
  return props
}

// ─── Styles, numbering, relationships ───────────────────────────────────

interface ThemeFonts {
  readonly major: string | undefined
  readonly minor: string | undefined
}

function readThemeFonts(xml: Element | undefined): ThemeFonts {
  const typeface = (kind: string): string | undefined => attr(kid(all(xml, `a:${kind}Font`)[0], 'a:latin'), 'typeface')
  return { major: typeface('major'), minor: typeface('minor') }
}

interface StyleDefinition {
  readonly type: string
  /** Lowercase, as Word matches built-in names. */
  readonly name: string
  readonly basedOn?: string
  readonly paragraph: ParagraphProps
  readonly run: RunProps
}

interface ResolvedStyle {
  readonly paragraph: ParagraphProps
  readonly run: RunProps
}

class StyleSheet {
  private readonly styles = new Map<string, StyleDefinition>()
  private readonly defaults: ResolvedStyle
  readonly defaultParagraph: string | undefined

  constructor(xml: Element | undefined, themeFonts: ThemeFonts) {
    const docDefaults = kid(xml, 'w:docDefaults')
    this.defaults = {
      paragraph: readParagraphProps(kid(kid(docDefaults, 'w:pPrDefault'), 'w:pPr')),
      run: readRunProps(kid(kid(docDefaults, 'w:rPrDefault'), 'w:rPr'), themeFonts),
    }
    let defaultParagraph: string | undefined
    for (const style of kids(xml, 'w:style')) {
      const id = attr(style, 'w:styleId')
      if (id === undefined) continue
      const type = attr(style, 'w:type') ?? 'paragraph'
      if (type === 'paragraph' && ['1', 'true', 'on'].includes(attr(style, 'w:default') ?? '')) defaultParagraph = id
      this.styles.set(id, {
        type,
        name: (val(style, 'w:name') ?? id).toLowerCase(),
        basedOn: val(style, 'w:basedOn'),
        paragraph: readParagraphProps(kid(style, 'w:pPr')),
        run: readRunProps(kid(style, 'w:rPr'), themeFonts),
      })
    }
    this.defaultParagraph = defaultParagraph
  }

  name(id: string | undefined): string | undefined {
    return id === undefined ? undefined : this.styles.get(id)?.name
  }

  /** A style with its `basedOn` chain applied, without the document defaults. */
  private chain(id: string | undefined, seen = new Set<string>()): ResolvedStyle {
    const style = id === undefined ? undefined : this.styles.get(id)
    if (id === undefined || style === undefined || seen.has(id)) return { paragraph: {}, run: {} }
    seen.add(id)
    const base = this.chain(style.basedOn, seen)
    return { paragraph: { ...base.paragraph, ...style.paragraph }, run: { ...base.run, ...style.run } }
  }

  /** A paragraph style over the document defaults. */
  paragraph(id: string | undefined): ResolvedStyle {
    const style = this.chain(id ?? this.defaultParagraph)
    return { paragraph: { ...this.defaults.paragraph, ...style.paragraph }, run: { ...this.defaults.run, ...style.run } }
  }

  character(id: string | undefined): RunProps {
    return this.chain(id).run
  }

  /** The list a numbering style (`w:numStyleLink`) stands for. */
  numberingStyleList(id: string): string | undefined {
    return this.chain(id).paragraph.numId
  }

  /** The id of the paragraph style with this (lowercase) name. */
  idOf(name: string): string | undefined {
    for (const [id, style] of this.styles) if (style.type === 'paragraph' && style.name === name) return id
    return undefined
  }
}

interface ListLevel {
  /** `bullet`, `none`, or a number format such as `decimal`. */
  readonly format: string
  readonly start: number
  /**
   * What the item counts in: lists sharing an abstract definition keep counting across each
   * other in Word, unless one of them restarts with a start override.
   */
  readonly counter: string
}

type Numbering = ReadonlyMap<string, ReadonlyMap<number, ListLevel>>

function readNumbering(xml: Element | undefined, styles: StyleSheet): Numbering {
  const abstract = new Map<string, { readonly levels: Map<number, Omit<ListLevel, 'counter'>>; readonly styleLink: string | undefined }>()
  for (const definition of kids(xml, 'w:abstractNum')) {
    const levels = new Map<number, Omit<ListLevel, 'counter'>>()
    for (const level of kids(definition, 'w:lvl')) {
      levels.set(number(attr(level, 'w:ilvl')) ?? 0, { format: val(level, 'w:numFmt') ?? 'decimal', start: number(val(level, 'w:start')) ?? 1 })
    }
    abstract.set(attr(definition, 'w:abstractNumId') ?? '', { levels, styleLink: val(definition, 'w:numStyleLink') })
  }
  const abstractOf = new Map(kids(xml, 'w:num').map(num => [attr(num, 'w:numId') ?? '', val(num, 'w:abstractNumId') ?? ''] as const))
  // A list defined through a list style points at that style, whose own list holds the levels.
  const levelsOf = (abstractId: string): { readonly id: string; readonly levels: Map<number, Omit<ListLevel, 'counter'>> } => {
    const definition = abstract.get(abstractId)
    const linked = definition?.styleLink === undefined ? undefined : styles.numberingStyleList(definition.styleLink)
    const target = linked === undefined ? undefined : abstractOf.get(linked)
    if (definition !== undefined && definition.levels.size === 0 && target !== undefined && target !== abstractId) {
      return { id: target, levels: abstract.get(target)?.levels ?? new Map() }
    }
    return { id: abstractId, levels: definition?.levels ?? new Map() }
  }
  const numbering = new Map<string, ReadonlyMap<number, ListLevel>>()
  for (const num of kids(xml, 'w:num')) {
    const numId = attr(num, 'w:numId') ?? ''
    const source = levelsOf(val(num, 'w:abstractNumId') ?? '')
    const levels = new Map<number, ListLevel>()
    for (const [index, level] of source.levels) levels.set(index, { ...level, counter: `abstract:${source.id}` })
    for (const override of kids(num, 'w:lvlOverride')) {
      const index = number(attr(override, 'w:ilvl')) ?? 0
      const start = number(val(override, 'w:startOverride'))
      const level = levels.get(index)
      if (level !== undefined && start !== undefined) levels.set(index, { ...level, start, counter: `num:${numId}` })
    }
    numbering.set(numId, levels)
  }
  return numbering
}

interface Relationship {
  readonly target: string
  readonly external: boolean
}

function readRelationships(xml: Element | undefined): ReadonlyMap<string, Relationship> {
  const relationships = new Map<string, Relationship>()
  for (const relationship of kids(xml, 'Relationship')) {
    const target = attr(relationship, 'Target') ?? ''
    const external = attr(relationship, 'TargetMode') === 'External'
    // Package targets are relative to word/, or absolute from the package root.
    const resolved = external ? target : target.startsWith('/') ? target.slice(1) : `word/${target}`.replace(/[^/]+\/\.\.\//g, '')
    relationships.set(attr(relationship, 'Id') ?? '', { target: resolved, external })
  }
  return relationships
}

function readProperties(xml: Element | undefined): DocumentProperties {
  const text = (name: string): string => all(xml, name)[0]?.textContent?.trim() ?? ''
  return { title: text('dc:title'), subject: text('dc:subject'), creator: text('dc:creator'), keywords: text('cp:keywords'), description: text('dc:description') }
}

// ─── Content ────────────────────────────────────────────────────────────

/** `cell`, `continue` and `checklist` are Sumi's own table-cell, list-continuation and checklist paragraphs. */
type BlockKind = StyleName | 'code' | 'rule' | 'checklist' | 'continue' | 'cell'

interface Field {
  instruction: string
  inResult: boolean
}

interface Context {
  readonly styles: StyleSheet
  readonly numbering: Numbering
  readonly relationships: ReadonlyMap<string, Relationship>
  readonly media: Readonly<Record<string, Uint8Array>>
  readonly themeFonts: ThemeFonts
  readonly theme: DocumentTheme
  readonly page: PageSetup
  readonly findings: Findings
  /** Open complex fields (`w:fldChar`); they can span runs and paragraphs. */
  readonly fields: Field[]
  /** Items numbered so far per counter and level, so a list that continues after a break keeps counting. */
  readonly listCounts: Map<string, number[]>
}

/** The kind of block a paragraph style makes, by the style's own name. */
const kindsByStyleName: Readonly<Record<string, BlockKind>> = {
  normal: 'normal',
  title: 'title',
  'heading 1': 'heading1',
  'heading 2': 'heading2',
  'heading 3': 'heading3',
  quote: 'quote',
  'intense quote': 'quote',
  [docxStyleNames.Code.toLowerCase()]: 'code',
  [docxStyleNames.HorizontalLine.toLowerCase()]: 'rule',
  [docxStyleNames.Checklist.toLowerCase()]: 'checklist',
  [docxStyleNames.ListContinue.toLowerCase()]: 'continue',
  [docxStyleNames.TableText.toLowerCase()]: 'cell',
  'list paragraph': 'normal',
  'no spacing': 'normal',
}

const deeperHeading = /^heading [4-9]$/

/** Elements inside a paragraph, a run or the body that are handled, or that carry nothing Sumi could show. */
const knownElements = new Set([
  'w:pPr', 'w:rPr', 'w:tcPr', 'w:trPr', 'w:tblPr', 'w:tblGrid', 'w:sectPr', 'w:sdtPr', 'w:sdtEndPr', 'w:smartTagPr', 'w:customXmlPr',
  'w:p', 'w:tbl', 'w:tr', 'w:tc', 'w:r', 'w:hyperlink', 'w:fldSimple', 'w:ins', 'w:del', 'w:moveTo', 'w:moveFrom',
  'w:smartTag', 'w:customXml', 'w:dir', 'w:bdo', 'w:sdt', 'w:sdtContent', 'm:oMath', 'm:oMathPara', 'w:altChunk',
  'w:bookmarkStart', 'w:bookmarkEnd', 'w:commentRangeStart', 'w:commentRangeEnd', 'w:proofErr', 'w:permStart', 'w:permEnd',
  'w:moveFromRangeStart', 'w:moveFromRangeEnd', 'w:moveToRangeStart', 'w:moveToRangeEnd',
  'w:customXmlInsRangeStart', 'w:customXmlInsRangeEnd', 'w:customXmlDelRangeStart', 'w:customXmlDelRangeEnd',
  'w:customXmlMoveFromRangeStart', 'w:customXmlMoveFromRangeEnd', 'w:customXmlMoveToRangeStart', 'w:customXmlMoveToRangeEnd',
  'w:t', 'w:tab', 'w:ptab', 'w:br', 'w:cr', 'w:noBreakHyphen', 'w:softHyphen', 'w:sym', 'w:drawing', 'w:pict', 'mc:AlternateContent',
  'w:object', 'w:footnoteReference', 'w:endnoteReference', 'w:commentReference', 'w:fldChar', 'w:instrText', 'w:delText',
  'w:delInstrText', 'w:lastRenderedPageBreak', 'w:ruby', 'w:annotationRef', 'w:footnoteRef', 'w:endnoteRef', 'w:separator',
  'w:continuationSeparator', 'w:fldData',
])

function reportUnknown(element: Element, context: Context): void {
  if (knownElements.has(element.tagName)) return
  context.findings.add('Other Word content', 'degraded', { name: element.tagName, alternative: 'Not shown, and not kept when saved.' })
}

type InlineSegment = JSONContent[]

function textNode(text: string, marks: JSONContent['marks']): JSONContent {
  return { type: 'text', text, ...(marks !== undefined && marks.length > 0 ? { marks } : {}) }
}

function fieldName(instruction: string): string {
  return instruction.trim().split(/\s+/)[0]?.toUpperCase() ?? ''
}

function hyperlinkTarget(instruction: string): string | undefined {
  return /^\s*HYPERLINK\s+"([^"]+)"/i.exec(instruction)?.[1]
}

/** Formatting set on the text, as marks. A link's own style gives its colour and underline, which the link mark shows. */
function marksFor(run: RunProps, href: string | undefined, code: boolean, context: Context): JSONContent['marks'] {
  if (code) return [{ type: 'code' }]
  const marks: NonNullable<JSONContent['marks']> = []
  if (run.bold === true) marks.push({ type: 'bold' })
  if (run.italic === true) marks.push({ type: 'italic' })
  if (run.underline !== undefined && run.underline !== 'none') {
    marks.push({ type: 'underline' })
    if (run.underline !== 'single' && run.underline !== 'words') {
      context.findings.add('Underline styles', 'degraded', { alternative: 'Shown as a single underline, and saved that way.' })
    }
  }
  if (run.strike === true) marks.push({ type: 'strike' })
  if (run.vertAlign === 'superscript') marks.push({ type: 'superscript' })
  if (run.vertAlign === 'subscript') marks.push({ type: 'subscript' })
  const style: Record<string, string> = {}
  if (run.font !== undefined) style['fontFamily'] = run.font
  if (run.size !== undefined) style['fontSize'] = `${run.size}pt`
  if (run.color !== undefined) style['color'] = run.color
  if (run.background !== undefined) style['backgroundColor'] = run.background
  if (Object.keys(style).length > 0) marks.push({ type: 'textStyle', attrs: style })
  if (href !== undefined) marks.push({ type: 'link', attrs: { href } })
  if (run.effects === true) context.findings.add('Character spacing and text effects', 'degraded', { alternative: 'Shown as plain text, and saved that way.' })
  return marks
}

const imageTypes: Readonly<Record<string, string>> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' }

/** The media parts pictures point at, so only those are unpacked. */
function imageTargets(relationships: ReadonlyMap<string, Relationship>): Set<string> {
  return new Set([...relationships.values()].filter(relationship => !relationship.external && /\.(png|jpe?g|gif|bmp)$/i.test(relationship.target)).map(relationship => relationship.target))
}

function imageFromDrawing(drawing: Element, context: Context): JSONContent | undefined {
  const blips = [...all(drawing, 'a:blip'), ...all(drawing, 'v:imagedata')]
  const shapes = all(drawing, 'w:txbxContent').length + all(drawing, 'wps:wsp').length + all(drawing, 'wpg:wgp').length
    + all(drawing, 'v:textbox').length + ['v:rect', 'v:oval', 'v:line', 'v:roundrect', 'v:polyline', 'v:arc'].reduce((sum, name) => sum + all(drawing, name).length, 0)
  if (shapes > 0) context.findings.add('Text boxes and shapes', 'dropped', { unit: 'object', alternative: 'Not shown, and not kept when saved, with any text inside them.' })
  if (blips.length > 1) context.findings.add('Grouped pictures', 'dropped', { unit: 'group', alternative: 'Only the first picture of each group is shown and kept.' })
  const blip = blips[0]
  if (blip === undefined) {
    if (all(drawing, 'c:chart').length > 0) context.findings.add('Charts', 'dropped', { unit: 'chart', alternative: notKept })
    else if (all(drawing, 'dgm:relIds').length > 0) context.findings.add('SmartArt', 'dropped', { unit: 'diagram', alternative: notKept })
    else if (shapes === 0) context.findings.add('Other Word content', 'degraded', { name: 'drawing', alternative: notKept })
    return undefined
  }
  const relationship = context.relationships.get(attr(blip, 'r:embed') ?? attr(blip, 'r:id') ?? attr(blip, 'r:link') ?? '')
  if (relationship?.external === true) {
    context.findings.add('Linked images', 'dropped', { unit: 'image', alternative: 'Not shown, and not kept when saved. Zendo only shows images stored in the document.' })
    return undefined
  }
  const extension = relationship?.target.split('.').pop()?.toLowerCase() ?? ''
  const mimeType = imageTypes[extension]
  const bytes = relationship === undefined ? undefined : context.media[relationship.target]
  if (mimeType === undefined && relationship !== undefined) {
    context.findings.add('Images in EMF, WMF or TIFF format', 'dropped', { unit: 'image', alternative: 'Not shown, and not kept when saved. Zendo cannot show these formats.' })
    return undefined
  }
  if (bytes === undefined || mimeType === undefined) {
    context.findings.add('Missing pictures', 'dropped', { unit: 'picture', alternative: 'The file refers to pictures it does not contain. Not kept when saved.' })
    return undefined
  }
  if (all(drawing, 'wp:anchor').length > 0) {
    context.findings.add('Image positions and text wrapping', 'degraded', { alternative: 'Shown in line with the text, and saved that way.' })
  }
  const extent = all(drawing, 'wp:extent')[0]
  const properties = all(drawing, 'wp:docPr')[0]
  const style = attr(all(drawing, 'v:shape')[0], 'style') ?? ''
  const pointSize = (name: string): number | undefined => number(new RegExp(`${name}:([\\d.]+)pt`).exec(style)?.[1])
  const vmlWidth = pointSize('width')
  const vmlHeight = pointSize('height')
  const width = extent === undefined ? vmlWidth === undefined ? undefined : vmlWidth * 4 / 3 : (number(attr(extent, 'cx')) ?? 0) / 9525
  const height = extent === undefined ? vmlHeight === undefined ? undefined : vmlHeight * 4 / 3 : (number(attr(extent, 'cy')) ?? 0) / 9525
  const alt = attr(properties, 'descr') || attr(properties, 'title')
  return {
    type: 'image',
    attrs: {
      src: toDataUrl(bytes, mimeType),
      ...(alt === undefined || alt === '' ? {} : { alt }),
      ...(width === undefined || height === undefined ? {} : { width: Math.round(width), height: Math.round(height) }),
    },
  }
}

const trackedChange = 'Shown with every change accepted, and saved that way. To review the changes, accept or reject them in Word first.'

/** Walks a paragraph's runs, hyperlinks and wrappers into text nodes, images and breaks, split where it has page breaks. */
function inlineContent(paragraph: Element, styleRun: RunProps, context: Context): InlineSegment[] {
  const segments: InlineSegment[] = [[]]
  const current = (): InlineSegment => segments[segments.length - 1] ?? []

  const run = (element: Element, href: string | undefined): void => {
    const rPr = kid(element, 'w:rPr')
    const styleId = val(rPr, 'w:rStyle')
    const styleName = context.styles.name(styleId)
    const code = styleName === docxStyleNames.InlineCode.toLowerCase()
    // The Hyperlink style's colour and underline are how Word shows a link; the link mark shows it here.
    const characterStyle = styleName === 'hyperlink' ? {} : context.styles.character(styleId)
    const direct = readRunProps(rPr, context.themeFonts)
    const props: RunProps = { ...styleRun, ...characterStyle, ...direct }
    if (props.hidden === true) {
      context.findings.add('Hidden text', 'dropped', { alternative: notKept })
      return
    }
    const fieldLink = context.fields.map(open => hyperlinkTarget(open.instruction)).find(target => target !== undefined)
    const link = href ?? safeHref(fieldLink)
    const marks = (): JSONContent['marks'] => marksFor(props, link, code, context)
    for (const child of Array.from(element.children)) {
      switch (child.tagName) {
        case 'w:fldChar': {
          const type = attr(child, 'w:fldCharType')
          const open = context.fields[context.fields.length - 1]
          if (type === 'begin') context.fields.push({ instruction: '', inResult: false })
          else if (type === 'separate' && open !== undefined) open.inResult = true
          else if (type === 'end') {
            const ended = context.fields.pop()
            const name = ended === undefined ? '' : fieldName(ended.instruction)
            if (name !== '' && name !== 'HYPERLINK') {
              context.findings.add('Fields', 'dropped', { name, alternative: 'Their last shown text is kept as plain text; they no longer update.' })
            }
          }
          break
        }
        case 'w:instrText': {
          const open = context.fields[context.fields.length - 1]
          if (open !== undefined) open.instruction += child.textContent ?? ''
          break
        }
        case 'w:ruby':
          // A phonetic guide over East Asian text: the base text is the text itself.
          context.findings.add('Phonetic guides', 'degraded', { alternative: 'The text is kept; the guides above it are not shown or kept.' })
          wrapped(kid(child, 'w:rubyBase') ?? child, 'w:r').forEach(base => run(base, href))
          break
        case 'w:rPr':
          break
        default:
          // Text between a field's start and its result is the field code, not document text.
          if (context.fields.some(open => !open.inResult)) break
          inlineChild(child, marks)
      }
    }
  }

  const inlineChild = (child: Element, marks: () => JSONContent['marks']): void => {
    switch (child.tagName) {
      case 'w:t':
        if ((child.textContent ?? '') !== '') current().push(textNode(child.textContent ?? '', marks()))
        return
      case 'w:tab':
      case 'w:ptab':
        current().push(textNode('\t', marks()))
        return
      case 'w:noBreakHyphen':
        current().push(textNode('‑', marks()))
        return
      case 'w:sym': {
        const code = Number.parseInt(attr(child, 'w:char') ?? '', 16)
        if (Number.isFinite(code)) current().push(textNode(String.fromCharCode(code), marks()))
        return
      }
      case 'w:br':
      case 'w:cr':
        if (attr(child, 'w:type') === 'page') segments.push([])
        else current().push({ type: 'hardBreak' })
        return
      case 'w:drawing':
      case 'w:pict': {
        const image = imageFromDrawing(child, context)
        if (image !== undefined) current().push(image)
        return
      }
      case 'mc:AlternateContent': {
        const choice = kid(child, 'mc:Choice') ?? kid(child, 'mc:Fallback')
        kids(choice).forEach(inner => inlineChild(inner, marks))
        return
      }
      case 'w:object':
        context.findings.add('Embedded objects', 'dropped', { unit: 'object', alternative: notKept })
        return
      case 'w:footnoteReference':
      case 'w:endnoteReference':
        context.findings.add('Footnotes and endnotes', 'dropped', { unit: 'note', alternative: 'Not shown, and not kept when saved, with their numbers in the text.' })
        return
      default:
        reportUnknown(child, context)
    }
  }

  const walk = (container: Element, href: string | undefined): void => {
    for (const child of Array.from(container.children)) {
      switch (child.tagName) {
        case 'w:r':
          run(child, href)
          break
        case 'w:hyperlink': {
          const relationship = context.relationships.get(attr(child, 'r:id') ?? '')
          const anchor = attr(child, 'w:anchor')
          const target = relationship?.external === true ? safeHref(`${relationship.target}${anchor === undefined ? '' : `#${anchor}`}`) : undefined
          if (relationship === undefined && anchor !== undefined) {
            context.findings.add('Links within the document', 'degraded', { unit: 'link', alternative: 'The link text is kept; the links are not.' })
          }
          walk(child, target)
          break
        }
        case 'w:fldSimple': {
          const instruction = attr(child, 'w:instr') ?? ''
          const name = fieldName(instruction)
          if (name !== 'HYPERLINK') context.findings.add('Fields', 'dropped', { name, alternative: 'Their last shown text is kept as plain text; they no longer update.' })
          walk(child, safeHref(hyperlinkTarget(instruction)) ?? href)
          break
        }
        case 'w:ins':
        case 'w:moveTo':
        case 'w:smartTag':
        case 'w:customXml':
        case 'w:dir':
        case 'w:bdo':
          walk(child, href)
          break
        case 'w:sdt':
          if (all(child, 'w14:checkbox').length > 0) {
            context.findings.add('Form fields and check boxes', 'dropped', { alternative: 'Their current text is kept as plain text.' })
          }
          walk(kid(child, 'w:sdtContent') ?? child, href)
          break
        case 'w:bookmarkStart':
          // Word adds hidden bookmarks (named with a leading underscore) for its own use.
          if (!(attr(child, 'w:name') ?? '_').startsWith('_')) {
            context.findings.add('Bookmarks', 'degraded', { unit: 'bookmark', alternative: 'Not kept when saved; links and cross-references to them stop working.' })
          }
          break
        case 'm:oMath':
        case 'm:oMathPara':
          context.findings.add('Equations', 'dropped', { unit: 'equation', alternative: notKept })
          break
        default:
          reportUnknown(child, context)
      }
    }
  }

  walk(paragraph, undefined)
  return segments
}

function blockKind(name: string, resolved: ParagraphProps, context: Context): { readonly kind: BlockKind; readonly otherStyle: boolean } {
  const kind = kindsByStyleName[name]
  if (kind !== undefined) return { kind, otherStyle: false }
  const level = deeperHeading.test(name) ? 3 : resolved.outlineLevel
  if (level !== undefined && level < 9) {
    if (level >= 3) context.findings.add('Headings below level 3', 'degraded', { alternative: 'Shown as Heading 3, and saved that way.' })
    return { kind: (['heading1', 'heading2', 'heading3'] as const)[Math.min(level, 2)] ?? 'heading3', otherStyle: level >= 3 }
  }
  context.findings.add('Other paragraph styles', 'degraded', {
    name: name.replace(/\b\w/g, letter => letter.toUpperCase()),
    alternative: 'Their formatting is kept on each paragraph; the style names are not.',
  })
  return { kind: 'normal', otherStyle: true }
}

/** The formatting a paragraph style adds to the theme style it is shown in. */
function styleDifference(style: ResolvedStyle, base: BlockStyle): { readonly paragraph: ParagraphProps; readonly run: RunProps } {
  const paragraph: ParagraphProps = {}
  if (style.paragraph.spaceBefore !== undefined && style.paragraph.spaceBefore !== base.spaceBefore) paragraph.spaceBefore = style.paragraph.spaceBefore
  if (style.paragraph.spaceAfter !== undefined && style.paragraph.spaceAfter !== base.spaceAfter) paragraph.spaceAfter = style.paragraph.spaceAfter
  if (style.paragraph.lineHeight !== undefined && style.paragraph.lineHeight !== base.lineHeight) paragraph.lineHeight = style.paragraph.lineHeight
  if (style.paragraph.indentLeft !== undefined && style.paragraph.indentLeft !== base.indentLeft) paragraph.indentLeft = style.paragraph.indentLeft
  assign(paragraph, 'indentFirstLine', style.paragraph.indentFirstLine)
  assign(paragraph, 'lineExact', style.paragraph.lineExact)
  const run: RunProps = {}
  if (style.run.font !== undefined && style.run.font !== base.fontFamily) run.font = style.run.font
  if (style.run.size !== undefined && style.run.size !== base.fontSize) run.size = style.run.size
  if (style.run.color !== undefined && style.run.color !== base.color) run.color = style.run.color
  if (style.run.bold === true && !base.bold) run.bold = true
  if (style.run.italic === true && !base.italic) run.italic = true
  for (const key of ['underline', 'strike', 'background', 'vertAlign', 'effects'] as const) assign(run, key, style.run[key])
  return { paragraph, run }
}

/** Paragraph attributes from the paragraph's own formatting. Numbered paragraphs take their indent from the list. */
function paragraphAttrs(props: ParagraphProps, align: string | undefined, fontSize: number, numbered: boolean, context: Context): Record<string, string | number> {
  const attrs: Record<string, string | number> = {}
  if (align !== undefined && align !== 'left') attrs['textAlign'] = align
  assign(attrs, 'spaceBefore', props.spaceBefore)
  assign(attrs, 'spaceAfter', props.spaceAfter)
  if (props.lineExact !== undefined) {
    context.findings.add('Exact line spacing', 'degraded', { alternative: 'Shown as the nearest multiple of single spacing, and saved that way.' })
    attrs['lineHeight'] = round(props.lineExact / (fontSize * singleLineHeight))
  } else assign(attrs, 'lineHeight', props.lineHeight)
  if (!numbered) {
    assign(attrs, 'indentLeft', props.indentLeft)
    assign(attrs, 'indentFirstLine', props.indentFirstLine)
  }
  return attrs
}

/**
 * Where a paragraph sits: in a quote, and inside a list item at `level`. Sumi's own files mark a
 * paragraph that continues a list item with a list level whose number format is `none`.
 */
interface Placement {
  readonly quote: boolean
  readonly level: number | undefined
}

/** One paragraph, before consecutive list, checklist, quote and code paragraphs are grouped. */
type Item =
  | { readonly kind: 'block'; readonly node: JSONContent; readonly place: Placement }
  | { readonly kind: 'code'; readonly text: string; readonly place: Placement }
  | {
    readonly kind: 'item'
    readonly list: 'bulletList' | 'orderedList' | 'taskList'
    readonly level: number
    /** Paragraphs of one list share it; a different list starts a new one. */
    readonly key: string
    readonly start: number
    readonly checked: boolean
    readonly node: JSONContent
    readonly quote: boolean
  }

const listFormats = ['decimal', 'lowerLetter', 'lowerRoman']

/** Counts an item in its list and level, clearing deeper levels as Word does; resolves the count before it. */
function countItem(context: Context, counter: string, level: number): number {
  const counts = context.listCounts.get(counter) ?? []
  const before = counts[level] ?? 0
  counts[level] = before + 1
  counts.length = level + 1
  context.listCounts.set(counter, counts)
  return before
}

function stripBox(content: JSONContent[]): { readonly checked: boolean; readonly content: JSONContent[] } {
  const [first, ...rest] = content
  const text = first?.text ?? ''
  if (first === undefined || !/^[☐☑] ?/.test(text)) return { checked: false, content }
  const remaining = text.replace(/^[☐☑] ?/, '')
  return { checked: text.startsWith(checklistMarks.done), content: [...(remaining === '' ? [] : [{ ...first, text: remaining }]), ...rest] }
}

function paragraphItems(paragraph: Element, context: Context): Item[] {
  const pPr = kid(paragraph, 'w:pPr')
  const styleId = val(pPr, 'w:pStyle')
  const style = context.styles.paragraph(styleId)
  const direct = readParagraphProps(pPr)
  const resolved: ParagraphProps = { ...style.paragraph, ...direct }
  const { kind, otherStyle } = blockKind(context.styles.name(styleId ?? context.styles.defaultParagraph) ?? 'normal', resolved, context)
  // Code and rule paragraphs get their shading and line from Sumi's own styles.
  if (direct.decorated === true || (style.paragraph.decorated === true && kind !== 'code' && kind !== 'rule')) {
    context.findings.add('Paragraph borders and shading', 'degraded', { alternative: 'Shown without them, and saved that way.' })
  }
  if (resolved.rightToLeft === true) context.findings.add('Right-to-left paragraphs', 'degraded', { alternative: 'Shown left to right, and saved that way.' })
  const themeStyle: StyleName = kind === 'title' || kind === 'heading1' || kind === 'heading2' || kind === 'heading3' || kind === 'quote' ? kind : 'normal'
  const base = context.theme[themeStyle]
  const extra = otherStyle ? styleDifference(style, base) : { paragraph: {}, run: {} }
  const segments = inlineContent(paragraph, { ...extra.run, hidden: style.run.hidden }, context)
  if (resolved.tabStops === true && segments.some(segment => segment.some(node => node.text?.includes('\t') === true))) {
    context.findings.add('Tab stops', 'degraded', { alternative: 'Tabs line up at the default positions, and save that way.' })
  }

  const levels = resolved.numId === undefined || resolved.numId === '0' ? undefined : context.numbering.get(resolved.numId)
  const level = resolved.level ?? 0
  const listLevel = levels?.get(level)
  const marker = listLevel?.format === 'none'
  const isList = listLevel !== undefined && !marker && (kind === 'normal' || kind === 'continue' || kind === 'quote')
  if (listLevel !== undefined && !marker && kind.startsWith('heading')) {
    context.findings.add('Numbered headings', 'degraded', { alternative: 'Shown without their numbers, and saved that way.' })
  }
  const place: Placement = {
    quote: kind === 'quote',
    // A continuation paragraph from before list levels were marked attaches to the deepest open list.
    level: marker ? level : kind === 'continue' ? Number.POSITIVE_INFINITY : undefined,
  }
  const props: ParagraphProps = { ...extra.paragraph, ...direct }
  const attrs = paragraphAttrs(props, resolved.align, style.run.size ?? base.fontSize, listLevel !== undefined, context)

  const items: Item[] = []
  const pageBreak = (): Item => ({ kind: 'block', node: { type: 'pageBreak' }, place: { quote: false, level: undefined } })
  if (resolved.pageBreakBefore === true) items.push(pageBreak())
  segments.forEach((content, index) => {
    if (index > 0) items.push(pageBreak())
    // A page break at the start or end of a paragraph leaves an empty part that Word never showed.
    if (segments.length > 1 && content.length === 0) return
    const node = (type: string, own: readonly JSONContent[] = content, extraAttrs: Record<string, string | number> = {}): JSONContent => {
      const allAttrs = { ...extraAttrs, ...attrs }
      return { type, ...(Object.keys(allAttrs).length > 0 ? { attrs: allAttrs } : {}), ...(own.length > 0 ? { content: [...own] } : {}) }
    }
    if (isList && listLevel !== undefined) {
      if (listLevel.format !== (listLevel.format === 'bullet' ? 'bullet' : listFormats[level % 3])) {
        context.findings.add('List numbering styles', 'degraded', { alternative: 'Shown as 1., a., i. and bullets, and saved that way.' })
      }
      const before = countItem(context, listLevel.counter, level)
      items.push({
        kind: 'item', list: listLevel.format === 'bullet' ? 'bulletList' : 'orderedList', level, key: resolved.numId ?? '',
        start: listLevel.start + before, checked: false, node: node('paragraph'), quote: kind === 'quote',
      })
      return
    }
    switch (kind) {
      case 'title':
        items.push({ kind: 'block', node: node('title'), place })
        return
      case 'heading1':
      case 'heading2':
      case 'heading3':
        items.push({ kind: 'block', node: node('heading', content, { level: Number(kind.slice(-1)) }), place })
        return
      case 'code':
        items.push({ kind: 'code', text: content.map(child => child.type === 'hardBreak' ? '\n' : child.text ?? '').join(''), place })
        return
      case 'rule':
        items.push({ kind: 'block', node: { type: 'horizontalRule' }, place })
        return
      case 'checklist': {
        const box = stripBox(content)
        items.push({ kind: 'item', list: 'taskList', level: marker ? level : 0, key: 'checklist', start: 1, checked: box.checked, node: node('paragraph', box.content), quote: false })
        return
      }
      default:
        items.push({ kind: 'block', node: node('paragraph'), place })
    }
  })
  return items
}

interface CellBuild {
  readonly header: boolean
  readonly colspan: number
  rowspan: number
  readonly widths: number[]
  readonly background: string | undefined
  readonly content: JSONContent[]
}

function tableBorders(table: Element): boolean {
  const unusual = (borders: Element | undefined): boolean => kids(borders).some(border => {
    const value = attr(border, 'w:val') ?? 'single'
    const color = (attr(border, 'w:color') ?? 'auto').toLowerCase()
    return value !== 'single' || !['auto', '000000'].includes(color) || (number(attr(border, 'w:sz')) ?? 4) > 8
  })
  const style = val(kid(table, 'w:tblPr'), 'w:tblStyle')
  return (style !== undefined && !['TableGrid', 'TableNormal'].includes(style))
    || unusual(kid(kid(table, 'w:tblPr'), 'w:tblBorders'))
    || all(table, 'w:tcBorders').some(unusual)
}

function table(element: Element, context: Context): JSONContent {
  const grid = kids(kid(element, 'w:tblGrid'), 'w:gridCol').map(column => Math.round((number(attr(column, 'w:w')) ?? 0) / 15))
  // Equal columns filling the text width are what a new table has: no widths of their own.
  const columnWidth = contentWidthPixels(context.page) / Math.max(1, grid.length)
  const plain = grid.length > 0 && grid.every(width => Math.abs(width - columnWidth) <= 1)
  if (tableBorders(element)) context.findings.add('Table borders and styles', 'degraded', { alternative: 'Shown with a thin border around every cell, and saved that way.' })
  // Cells merged down (vMerge) grow the cell that starts the merge, by column.
  const merging = new Map<number, CellBuild>()
  const rows: CellBuild[][] = []
  for (const row of wrapped(element, 'w:tr')) {
    const trPr = kid(row, 'w:trPr')
    // A row deleted with tracked changes on is left out, as accepting the change would.
    if (kid(trPr, 'w:del') !== undefined) continue
    const header = toggle(trPr, 'w:tblHeader') === true
    let column = number(val(trPr, 'w:gridBefore')) ?? 0
    const cells: CellBuild[] = []
    for (const cell of wrapped(row, 'w:tc')) {
      const tcPr = kid(cell, 'w:tcPr')
      const span = number(val(tcPr, 'w:gridSpan')) ?? 1
      const vMerge = kid(tcPr, 'w:vMerge')
      const origin = merging.get(column)
      if (vMerge !== undefined && attr(vMerge, 'w:val') !== 'restart' && origin !== undefined) {
        origin.rowspan += 1
        column += span
        continue
      }
      const fill = attr(kid(tcPr, 'w:shd'), 'w:fill')
      const content = blocks(cell, context)
      const build: CellBuild = {
        header,
        colspan: span,
        rowspan: 1,
        widths: grid.slice(column, column + span),
        background: fill === undefined || fill === 'auto' ? undefined : `#${fill.toLowerCase()}`,
        content: content.length > 0 ? content : [{ type: 'paragraph' }],
      }
      if (vMerge !== undefined) merging.set(column, build)
      else merging.delete(column)
      cells.push(build)
      column += span
    }
    if (cells.length > 0) rows.push(cells)
  }
  return {
    type: 'table',
    content: rows.map(cells => ({
      type: 'tableRow',
      content: cells.map(cell => ({
        type: cell.header ? 'tableHeader' : 'tableCell',
        attrs: {
          colspan: cell.colspan,
          rowspan: cell.rowspan,
          colwidth: !plain && cell.widths.length === cell.colspan && cell.widths.every(width => width > 0) ? cell.widths : null,
          ...(cell.background === undefined ? {} : { backgroundColor: cell.background }),
        },
        content: cell.content,
      })),
    })),
  }
}

/** Builds the lists: items join the open list of their level and kind, and blocks placed at a level join its last item. */
function nest(items: readonly Item[]): JSONContent[] {
  const result: JSONContent[] = []
  let open: Array<{ readonly list: JSONContent; readonly level: number; readonly type: string; readonly key: string }> = []
  const lastItem = (list: JSONContent): JSONContent | undefined => list.content?.[list.content.length - 1]

  for (const item of items) {
    if (item.kind === 'item') {
      while (open.length > 0 && (open[open.length - 1]?.level ?? 0) > item.level) open.pop()
      const top = open[open.length - 1]
      if (top !== undefined && top.level === item.level && (top.type !== item.list || top.key !== item.key)) open.pop()
      const parent = open[open.length - 1]
      if (parent === undefined || parent.level < item.level) {
        const list: JSONContent = { type: item.list, content: [], ...(item.list === 'orderedList' && item.start !== 1 ? { attrs: { start: item.start } } : {}) }
        const holder = parent === undefined ? undefined : lastItem(parent.list)
        if (holder !== undefined) holder.content?.push(list)
        else result.push(list)
        open.push({ list, level: item.level, type: item.list, key: item.key })
      }
      const entry = item.list === 'taskList'
        ? { type: 'taskItem', attrs: { checked: item.checked }, content: [item.node] }
        : { type: 'listItem', content: [item.node] }
      open[open.length - 1]?.list.content?.push(entry)
      continue
    }
    const node = item.kind === 'code'
      ? { type: 'codeBlock', ...(item.text === '' ? {} : { content: [{ type: 'text', text: item.text }] }) }
      : item.node
    const level = item.place.level
    if (level === undefined) {
      open = []
      result.push(node)
      continue
    }
    while (open.length > 0 && (open[open.length - 1]?.level ?? 0) > level) open.pop()
    const top = open[open.length - 1]
    const holder = top === undefined ? undefined : lastItem(top.list)
    if (holder !== undefined) holder.content?.push(node)
    else result.push(node)
  }
  return result
}

/** Joins code lines into code blocks, then puts quoted runs into quotations. */
function group(items: readonly Item[]): JSONContent[] {
  const joined: Item[] = []
  for (const item of items) {
    const previous = joined[joined.length - 1]
    if (item.kind === 'code' && previous?.kind === 'code' && previous.place.quote === item.place.quote && previous.place.level === item.place.level) {
      joined[joined.length - 1] = { ...previous, text: `${previous.text}\n${item.text}` }
    } else joined.push(item)
  }
  const quoted = (item: Item): boolean => item.kind === 'item' ? item.quote : item.place.quote
  const result: JSONContent[] = []
  let run: Item[] = []
  let runQuoted = false
  const flush = (): void => {
    if (run.length > 0) result.push(...(runQuoted ? [{ type: 'blockquote', content: nest(run) }] : nest(run)))
    run = []
  }
  for (const item of joined) {
    if (run.length > 0 && quoted(item) !== runQuoted) flush()
    runQuoted = quoted(item)
    run.push(item)
  }
  flush()
  return result
}

function blocks(container: Element, context: Context): JSONContent[] {
  const items: Item[] = []
  const collect = (parent: Element): void => {
    for (const child of Array.from(parent.children)) {
      switch (child.tagName) {
        case 'w:p':
          items.push(...paragraphItems(child, context))
          break
        case 'w:tbl':
          items.push({ kind: 'block', node: table(child, context), place: { quote: false, level: undefined } })
          break
        case 'w:sdt':
          collect(kid(child, 'w:sdtContent') ?? child)
          break
        case 'w:customXml':
        case 'w:ins':
        case 'w:moveTo':
          collect(child)
          break
        case 'm:oMathPara':
          context.findings.add('Equations', 'dropped', { unit: 'equation', alternative: notKept })
          break
        case 'w:altChunk':
          context.findings.add('Embedded documents', 'dropped', { alternative: notKept })
          break
        case 'w:bookmarkStart':
          if (!(attr(child, 'w:name') ?? '_').startsWith('_')) {
            context.findings.add('Bookmarks', 'degraded', { unit: 'bookmark', alternative: 'Not kept when saved; links and cross-references to them stop working.' })
          }
          break
        default:
          reportUnknown(child, context)
      }
    }
  }
  collect(container)
  return group(items)
}

// ─── Theme, page setup ──────────────────────────────────────────────────

function blockStyle(style: ResolvedStyle, fallback: BlockStyle): BlockStyle {
  return {
    fontFamily: style.run.font ?? fallback.fontFamily,
    fontSize: style.run.size ?? fallback.fontSize,
    color: style.run.color ?? fallback.color,
    bold: style.run.bold ?? false,
    italic: style.run.italic ?? false,
    spaceBefore: style.paragraph.spaceBefore ?? 0,
    spaceAfter: style.paragraph.spaceAfter ?? 0,
    lineHeight: style.paragraph.lineHeight ?? 1,
    indentLeft: style.paragraph.indentLeft ?? 0,
  }
}

const themeStyleNames: Readonly<Record<Exclude<StyleName, 'normal'>, readonly string[]>> = {
  title: ['title'],
  heading1: ['heading 1'],
  heading2: ['heading 2'],
  heading3: ['heading 3'],
  quote: ['quote', 'intense quote'],
}

/** The document's own look: its styles where it defines them, otherwise Sumi's in its body font. */
function readTheme(styles: StyleSheet): DocumentTheme {
  // Word shows text 10 pt Times New Roman when a file sets no default at all.
  const normal = blockStyle(styles.paragraph(undefined), { ...defaultTheme.normal, fontFamily: 'Times New Roman', fontSize: 10 })
  const style = (name: Exclude<StyleName, 'normal'>): BlockStyle => {
    const id = themeStyleNames[name].map(styleName => styles.idOf(styleName)).find(found => found !== undefined)
    const fallback = { ...defaultTheme[name], fontFamily: normal.fontFamily }
    return id === undefined ? fallback : blockStyle(styles.paragraph(id), fallback)
  }
  return { normal, title: style('title'), heading1: style('heading1'), heading2: style('heading2'), heading3: style('heading3'), quote: style('quote') }
}

/** Whether a header or footer shows anything besides a page number, and whether it has one. */
function readHeaderPart(xml: Element | undefined): { readonly content: boolean; readonly pageNumber: boolean } {
  if (xml === undefined) return { content: false, pageNumber: false }
  const fields = [...all(xml, 'w:instrText').map(element => element.textContent ?? ''), ...all(xml, 'w:fldSimple').map(element => attr(element, 'w:instr') ?? '')]
    .map(fieldName).filter(name => name !== '')
  const pageNumber = fields.includes('PAGE')
  let depth = 0
  let inResult = false
  // Any field other than the page number (a date, the file name) is content Sumi does not keep.
  let content = fields.some(name => name !== 'PAGE') || all(xml, 'w:drawing').length > 0 || all(xml, 'w:pict').length > 0
  // Runs directly in a w:fldSimple are its result, not text of their own.
  for (const run of all(xml, 'w:r').filter(candidate => candidate.parentElement?.tagName !== 'w:fldSimple')) {
    for (const child of Array.from(run.children)) {
      if (child.tagName === 'w:fldChar') {
        const type = attr(child, 'w:fldCharType')
        if (type === 'begin') depth += 1
        if (type === 'separate') inResult = true
        if (type === 'end') {
          depth -= 1
          inResult = false
        }
      } else if (child.tagName === 'w:t' && depth === 0 && !inResult && (child.textContent ?? '').trim() !== '') content = true
    }
  }
  return { content, pageNumber }
}

/**
 * The page setup of the last section, which Sumi keeps. Every section is checked for headers,
 * footers and columns, since an earlier section can have them when the last one does not.
 */
function readSections(body: Element, relationships: ReadonlyMap<string, Relationship>, parts: Readonly<Record<string, Uint8Array>>, findings: Findings): PageSetup {
  const sections = all(body, 'w:sectPr').filter(section => section.parentElement?.tagName !== 'w:sectPrChange')
  const last = kid(body, 'w:sectPr') ?? sections[sections.length - 1]
  let pageNumbers = false
  let headerContent = false
  for (const section of sections) {
    if ((number(attr(kid(section, 'w:cols'), 'w:num')) ?? 1) > 1) {
      findings.add('Multiple columns', 'dropped', { alternative: 'Shown in one column, and saved that way.' })
    }
    for (const reference of [...kids(section, 'w:headerReference'), ...kids(section, 'w:footerReference')]) {
      const target = relationships.get(attr(reference, 'r:id') ?? '')?.target
      const part = target === undefined ? undefined : parts[target]
      const read = readHeaderPart(part === undefined ? undefined : parseXml(strFromU8(part)))
      const footer = reference.tagName === 'w:footerReference'
      if (footer && read.pageNumber && section === last) pageNumbers = true
      if (read.content || (read.pageNumber && (!footer || section !== last))) headerContent = true
    }
  }
  if (headerContent) {
    findings.add('Header and footer text', 'dropped', {
      alternative: pageNumbers ? 'Not shown, and not kept when saved. Page numbers are kept.' : notKept,
    })
  }
  if (sections.length > 1) {
    findings.add('Section breaks', 'dropped', { unit: 'section', count: sections.length, alternative: 'Saved as one section, with the last section’s page size and margins.' })
  }
  if (last === undefined) return defaultPage
  const size = kid(last, 'w:pgSz')
  const margins = kid(last, 'w:pgMar')
  const length = (element: Element | undefined, name: string, fallback: number): number => Math.abs(number(attr(element, name)) ?? fallback)
  return {
    width: length(size, 'w:w', defaultPage.width),
    height: length(size, 'w:h', defaultPage.height),
    marginTop: length(margins, 'w:top', defaultPage.marginTop),
    marginRight: length(margins, 'w:right', defaultPage.marginRight),
    marginBottom: length(margins, 'w:bottom', defaultPage.marginBottom),
    marginLeft: length(margins, 'w:left', defaultPage.marginLeft),
    pageNumbers,
  }
}

/** Tracked changes anywhere in the body: insertions and deletions, and formatting changes. */
function trackedChanges(body: Element): number {
  return ['w:ins', 'w:del', 'w:moveFrom', 'w:moveTo', 'w:rPrChange', 'w:pPrChange', 'w:sectPrChange', 'w:tblPrChange', 'w:trPrChange', 'w:tcPrChange', 'w:numberingChange']
    .reduce((sum, name) => sum + all(body, name).length, 0)
}

function readSettings(xml: Element | undefined, findings: Findings): void {
  if (kid(xml, 'w:documentProtection') !== undefined) {
    findings.add('Editing restrictions', 'dropped', { alternative: 'Not kept when saved, so the saved file can be edited freely.' })
  }
  if (toggle(xml, 'w:trackRevisions') === true) {
    findings.add('Track changes setting', 'degraded', { alternative: 'Word stops tracking changes in the saved file until you turn it on again.' })
  }
}

export function readDocx(bytes: Uint8Array): DocxRead {
  assertZipFitsInMemory(bytes)
  const parts = unzipParts(bytes, name => xmlParts.has(name) || headerOrFooter.test(name))
  const xmlPart = (name: string): Element | undefined => {
    const part = parts[name]
    return part === undefined ? undefined : parseXml(strFromU8(part))
  }
  const body = kid(xmlPart('word/document.xml'), 'w:body')
  if (body === undefined) throw new Error('This file is not a Word document, or it is damaged.')

  const relationships = readRelationships(xmlPart('word/_rels/document.xml.rels'))
  const pictures = imageTargets(relationships)
  const media = pictures.size === 0 ? {} : unzipParts(bytes, name => pictures.has(name))
  const themeFonts = readThemeFonts(xmlPart('word/theme/theme1.xml'))
  const styles = new StyleSheet(xmlPart('word/styles.xml'), themeFonts)
  const findings = new Findings()
  const page = readSections(body, relationships, parts, findings)
  const context: Context = {
    styles,
    numbering: readNumbering(xmlPart('word/numbering.xml'), styles),
    relationships,
    media,
    themeFonts,
    theme: readTheme(styles),
    page,
    findings,
    fields: [],
    listCounts: new Map(),
  }

  const comments = kids(xmlPart('word/comments.xml'), 'w:comment').length
  if (comments > 0) findings.add('Comments', 'dropped', { unit: 'comment', count: comments, alternative: notKept })
  const changes = trackedChanges(body)
  if (changes > 0) findings.add('Tracked changes', 'dropped', { unit: 'change', count: changes, alternative: trackedChange })
  readSettings(xmlPart('word/settings.xml'), findings)

  const content = blocks(body, context)
  return {
    content: {
      type: 'doc',
      attrs: { page, theme: context.theme, properties: readProperties(xmlPart('docProps/core.xml')) },
      content: content.length > 0 ? content : [{ type: 'paragraph' }],
    },
    findings: findings.list(),
  }
}
