/**
 * Reads .docx into a Sumi document. The package is unzipped with fflate and its XML parsed with
 * DOMParser, then mapped straight to TipTap JSON:
 * - styles.xml becomes the document's theme (Normal, Title, Heading 1–3, Quote); any other style
 *   and all direct formatting become formatting on the paragraphs and text themselves.
 * - numbering.xml says which paragraphs are bulleted or numbered, and at what level.
 * - document.xml gives the content; its last section gives the page setup; a PAGE field in the
 *   footer turns page numbers on.
 * Whatever is not kept is listed in the import report, with what happens to it on save.
 */

import type { JSONContent } from '@tiptap/core'
import { strFromU8, unzipSync } from 'fflate'
import type { FindingSeverity, ImportFindingInput } from '@shared/fidelity'
import { defaultPage, defaultTheme, singleLineHeight, type BlockStyle, type DocumentTheme, type PageSetup, type StyleName } from '../theme'
import { checklistMarks, docxStyleNames, highlightColors } from './docx-styles'
import { toDataUrl } from './images'

export interface DocxRead {
  readonly content: JSONContent
  readonly findings: readonly ImportFindingInput[]
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

function all(element: Element, name: string): Element[] {
  return Array.from(element.getElementsByTagName(name))
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

const removed = 'Removed when saved; the original file keeps them.'

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
  const strike = toggle(rPr, 'w:strike') ?? toggle(rPr, 'w:dstrike')
  const effects = ['w:caps', 'w:smallCaps', 'w:shadow', 'w:outline', 'w:emboss', 'w:imprint'].some(name => toggle(rPr, name) === true)
    || kid(rPr, 'w:spacing') !== undefined || kid(rPr, 'w:w') !== undefined || kid(rPr, 'w:position') !== undefined
  const props: RunProps = {}
  assign(props, 'bold', toggle(rPr, 'w:b'))
  assign(props, 'italic', toggle(rPr, 'w:i'))
  assign(props, 'underline', val(rPr, 'w:u') ?? (kid(rPr, 'w:u') === undefined ? undefined : 'single'))
  assign(props, 'strike', strike)
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
  const typeface = (kind: string): string | undefined => {
    const font = xml === undefined ? undefined : all(xml, `a:${kind}Font`)[0]
    return attr(kid(font, 'a:latin'), 'typeface')
  }
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
    const docDefaults = xml === undefined ? undefined : kid(xml, 'w:docDefaults')
    this.defaults = {
      paragraph: readParagraphProps(kid(kid(docDefaults, 'w:pPrDefault'), 'w:pPr')),
      run: readRunProps(kid(kid(docDefaults, 'w:rPrDefault'), 'w:rPr'), themeFonts),
    }
    let defaultParagraph: string | undefined
    for (const style of kids(xml, 'w:style')) {
      const id = attr(style, 'w:styleId')
      if (id === undefined) continue
      const type = attr(style, 'w:type') ?? 'paragraph'
      if (type === 'paragraph' && toggle(style, 'w:default') === undefined && attr(style, 'w:default') === '1') defaultParagraph = id
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

  /** The id of the paragraph style with this (lowercase) name. */
  idOf(name: string): string | undefined {
    for (const [id, style] of this.styles) if (style.type === 'paragraph' && style.name === name) return id
    return undefined
  }
}

interface ListLevel {
  readonly ordered: boolean
  readonly format: string
  readonly start: number
}

type Numbering = ReadonlyMap<string, ReadonlyMap<number, ListLevel>>

function readNumbering(xml: Element | undefined): Numbering {
  const abstract = new Map<string, Map<number, ListLevel>>()
  for (const definition of kids(xml, 'w:abstractNum')) {
    const levels = new Map<number, ListLevel>()
    for (const level of kids(definition, 'w:lvl')) {
      const format = val(level, 'w:numFmt') ?? 'decimal'
      levels.set(number(attr(level, 'w:ilvl')) ?? 0, { ordered: format !== 'bullet' && format !== 'none', format, start: number(val(level, 'w:start')) ?? 1 })
    }
    abstract.set(attr(definition, 'w:abstractNumId') ?? '', levels)
  }
  const numbering = new Map<string, ReadonlyMap<number, ListLevel>>()
  for (const num of kids(xml, 'w:num')) {
    const levels = new Map(abstract.get(val(num, 'w:abstractNumId') ?? '') ?? [])
    for (const override of kids(num, 'w:lvlOverride')) {
      const index = number(attr(override, 'w:ilvl')) ?? 0
      const start = number(val(override, 'w:startOverride'))
      const level = levels.get(index)
      if (level !== undefined && start !== undefined) levels.set(index, { ...level, start })
    }
    numbering.set(attr(num, 'w:numId') ?? '', levels)
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

// ─── Content ────────────────────────────────────────────────────────────

/** `cell` and `continue` are Sumi's table-cell and list-continuation paragraphs, whose spacing and indent come from their style. */
type BlockKind = StyleName | 'code' | 'rule' | 'checklist' | 'continue' | 'cell'

interface Field {
  instruction: string
  inResult: boolean
}

interface Context {
  readonly styles: StyleSheet
  readonly numbering: Numbering
  readonly relationships: ReadonlyMap<string, Relationship>
  readonly parts: Readonly<Record<string, Uint8Array>>
  readonly themeFonts: ThemeFonts
  readonly theme: DocumentTheme
  readonly findings: Findings
  /** Open complex fields (`w:fldChar`); they can span runs and paragraphs. */
  readonly fields: Field[]
  /** Items numbered so far per list and level, so a list that continues after a break keeps counting. */
  readonly listCounts: Map<string, number[]>
  sections: number
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

type InlineSegment = JSONContent[]

function textNode(text: string, marks: JSONContent['marks']): JSONContent {
  return { type: 'text', text, ...(marks !== undefined && marks.length > 0 ? { marks } : {}) }
}

function safeHref(href: string | undefined): string | undefined {
  return href !== undefined && /^(https?:|mailto:)/i.test(href.trim()) ? href.trim() : undefined
}

function fieldName(instruction: string): string {
  return instruction.trim().split(/\s+/)[0]?.toUpperCase() ?? ''
}

function hyperlinkTarget(instruction: string): string | undefined {
  return /^\s*HYPERLINK\s+"([^"]+)"/i.exec(instruction)?.[1]
}

function marksFor(run: RunProps, base: BlockStyle, href: string | undefined, code: boolean, context: Context): JSONContent['marks'] {
  if (code) return [{ type: 'code' }]
  const marks: NonNullable<JSONContent['marks']> = []
  if (run.bold === true && !base.bold) marks.push({ type: 'bold' })
  if (run.italic === true && !base.italic) marks.push({ type: 'italic' })
  // A link's own style underlines and colours it; the link mark shows that.
  if (href === undefined && run.underline !== undefined && run.underline !== 'none') {
    marks.push({ type: 'underline' })
    if (run.underline !== 'single' && run.underline !== 'words') {
      context.findings.add('Underline styles', 'degraded', { alternative: 'Shown and saved as a single underline.' })
    }
  }
  if (run.strike === true) marks.push({ type: 'strike' })
  if (run.vertAlign === 'superscript') marks.push({ type: 'superscript' })
  if (run.vertAlign === 'subscript') marks.push({ type: 'subscript' })
  const style: Record<string, string> = {}
  if (run.font !== undefined && run.font !== base.fontFamily) style['fontFamily'] = run.font
  if (run.size !== undefined && run.size !== base.fontSize) style['fontSize'] = `${run.size}pt`
  if (href === undefined && run.color !== undefined && run.color !== base.color) style['color'] = run.color
  if (run.background !== undefined) style['backgroundColor'] = run.background
  if (Object.keys(style).length > 0) marks.push({ type: 'textStyle', attrs: style })
  if (href !== undefined) marks.push({ type: 'link', attrs: { href } })
  if (run.effects === true) context.findings.add('Character spacing and text effects', 'degraded', { alternative: 'Shown and saved as plain text.' })
  return marks
}

function imageFromDrawing(drawing: Element, context: Context): JSONContent | undefined {
  const blip = all(drawing, 'a:blip')[0] ?? all(drawing, 'v:imagedata')[0]
  if (blip === undefined) {
    if (all(drawing, 'c:chart').length > 0) context.findings.add('Charts', 'dropped', { unit: 'chart', alternative: removed })
    else if (all(drawing, 'dgm:relIds').length > 0) context.findings.add('SmartArt', 'dropped', { unit: 'diagram', alternative: removed })
    else context.findings.add('Text boxes and shapes', 'dropped', { unit: 'object', alternative: 'Removed when saved, with any text inside them; the original file keeps them.' })
    return undefined
  }
  const relationship = context.relationships.get(attr(blip, 'r:embed') ?? attr(blip, 'r:id') ?? attr(blip, 'r:link') ?? '')
  if (relationship?.external === true) {
    context.findings.add('Linked images', 'dropped', { unit: 'image', alternative: 'Removed when saved; Zendo only shows images stored in the document.' })
    return undefined
  }
  const bytes = relationship === undefined ? undefined : context.parts[relationship.target]
  const extension = relationship?.target.split('.').pop()?.toLowerCase() ?? ''
  const mimeType = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' }[extension]
  if (bytes === undefined || mimeType === undefined) {
    context.findings.add('Images in EMF, WMF or TIFF format', 'dropped', { unit: 'image', alternative: 'Removed when saved; Zendo cannot show these formats.' })
    return undefined
  }
  if (all(drawing, 'wp:anchor').length > 0) {
    context.findings.add('Image positions and text wrapping', 'degraded', { alternative: 'Shown and saved in line with the text.' })
  }
  const extent = all(drawing, 'wp:extent')[0]
  const properties = all(drawing, 'wp:docPr')[0]
  const style = attr(all(drawing, 'v:shape')[0], 'style') ?? ''
  const pointSize = (name: string): number | undefined => number(new RegExp(`${name}:([\\d.]+)pt`).exec(style)?.[1])
  const width = extent === undefined ? pointSize('width') === undefined ? undefined : (pointSize('width') ?? 0) * 4 / 3 : (number(attr(extent, 'cx')) ?? 0) / 9525
  const height = extent === undefined ? pointSize('height') === undefined ? undefined : (pointSize('height') ?? 0) * 4 / 3 : (number(attr(extent, 'cy')) ?? 0) / 9525
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

/** Walks a paragraph's runs, hyperlinks and wrappers into text nodes, images and breaks, split where it has page breaks. */
function inlineContent(paragraph: Element, base: BlockStyle, paragraphRun: RunProps, context: Context): InlineSegment[] {
  const segments: InlineSegment[] = [[]]
  const current = (): InlineSegment => segments[segments.length - 1] ?? []

  const run = (element: Element, href: string | undefined): void => {
    const rPr = kid(element, 'w:rPr')
    const styleId = val(rPr, 'w:rStyle')
    const code = context.styles.name(styleId) === docxStyleNames.InlineCode.toLowerCase()
    const props: RunProps = { ...paragraphRun, ...context.styles.character(styleId), ...readRunProps(rPr, context.themeFonts) }
    if (props.hidden === true) {
      context.findings.add('Hidden text', 'dropped', { alternative: 'Removed when saved; the original file keeps it.' })
      return
    }
    const fieldLink = context.fields.map(open => hyperlinkTarget(open.instruction)).find(target => target !== undefined)
    const link = href ?? safeHref(fieldLink)
    const marks = (): JSONContent['marks'] => marksFor(props, base, link, code, context)
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
              context.findings.add('Fields', 'dropped', { name, alternative: 'Kept as their last shown text; they no longer update.' })
            }
          }
          break
        }
        case 'w:instrText': {
          const open = context.fields[context.fields.length - 1]
          if (open !== undefined) open.instruction += child.textContent ?? ''
          break
        }
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
        context.findings.add('Embedded objects', 'dropped', { unit: 'object', alternative: removed })
        return
      case 'w:footnoteReference':
      case 'w:endnoteReference':
        context.findings.add('Footnotes and endnotes', 'dropped', { unit: 'note', alternative: 'Removed when saved, with their numbers in the text; the original file keeps them.' })
        return
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
          walk(child, relationship?.external === true ? safeHref(relationship.target) : undefined)
          break
        }
        case 'w:fldSimple': {
          const instruction = attr(child, 'w:instr') ?? ''
          const name = fieldName(instruction)
          if (name !== 'HYPERLINK') context.findings.add('Fields', 'dropped', { name, alternative: 'Kept as their last shown text; they no longer update.' })
          walk(child, safeHref(hyperlinkTarget(instruction)) ?? href)
          break
        }
        case 'w:ins':
        case 'w:moveTo':
          context.findings.add('Tracked changes', 'dropped', { unit: 'change', alternative: 'Shown with every change accepted. To review them, accept or reject the changes in Word first.' })
          walk(child, href)
          break
        case 'w:del':
        case 'w:moveFrom':
          context.findings.add('Tracked changes', 'dropped', { unit: 'change', alternative: 'Shown with every change accepted. To review them, accept or reject the changes in Word first.' })
          break
        case 'w:smartTag':
        case 'w:customXml':
        case 'w:dir':
        case 'w:bdo':
          walk(child, href)
          break
        case 'w:sdt':
          if (all(child, 'w14:checkbox').length > 0) {
            context.findings.add('Form fields and check boxes', 'dropped', { alternative: 'Kept as plain text.' })
          }
          walk(kid(child, 'w:sdtContent') ?? child, href)
          break
        case 'm:oMath':
        case 'm:oMathPara':
          context.findings.add('Equations', 'dropped', { unit: 'equation', alternative: removed })
          break
      }
    }
  }

  walk(paragraph, undefined)
  return segments
}

function blockKind(styleId: string | undefined, resolved: ParagraphProps, context: Context): BlockKind {
  const name = context.styles.name(styleId ?? context.styles.defaultParagraph) ?? 'normal'
  const kind = kindsByStyleName[name]
  if (kind !== undefined) return kind
  const level = deeperHeading.test(name) ? 3 : resolved.outlineLevel
  if (level !== undefined && level < 9) {
    if (level >= 3) context.findings.add('Headings below level 3', 'degraded', { alternative: 'Shown and saved as Heading 3.' })
    return (['heading1', 'heading2', 'heading3'] as const)[Math.min(level, 2)] ?? 'heading3'
  }
  context.findings.add('Other paragraph styles', 'degraded', {
    name: name.replace(/\b\w/g, letter => letter.toUpperCase()),
    alternative: 'Their formatting is kept on each paragraph; the style names are not.',
  })
  return 'normal'
}

function paragraphAttrs(props: ParagraphProps, base: BlockStyle, fontSize: number, spacing: boolean, context: Context): Record<string, string | number> {
  const attrs: Record<string, string | number> = {}
  if (props.align !== undefined && props.align !== 'left') attrs['textAlign'] = props.align
  if (!spacing) return attrs
  if (props.spaceBefore !== undefined && props.spaceBefore !== base.spaceBefore) attrs['spaceBefore'] = props.spaceBefore
  if (props.spaceAfter !== undefined && props.spaceAfter !== base.spaceAfter) attrs['spaceAfter'] = props.spaceAfter
  if (props.lineExact !== undefined) {
    context.findings.add('Exact line spacing', 'degraded', { alternative: 'Shown and saved as the nearest multiple of single spacing.' })
    attrs['lineHeight'] = round(props.lineExact / (fontSize * singleLineHeight))
  } else if (props.lineHeight !== undefined && props.lineHeight !== base.lineHeight) attrs['lineHeight'] = props.lineHeight
  if (props.indentLeft !== undefined && props.indentLeft !== base.indentLeft) attrs['indentLeft'] = props.indentLeft
  if (props.indentFirstLine !== undefined && props.indentFirstLine !== 0) attrs['indentFirstLine'] = props.indentFirstLine
  return attrs
}

/** One paragraph, before consecutive list, quote, code and checklist paragraphs are grouped. */
type Item =
  | { readonly kind: 'block'; readonly node: JSONContent }
  | { readonly kind: 'list'; readonly numId: string; readonly level: number; readonly ordered: boolean; readonly start: number; readonly node: JSONContent }
  | { readonly kind: 'checklist'; readonly checked: boolean; readonly node: JSONContent }
  | { readonly kind: 'code'; readonly text: string }
  | { readonly kind: 'quote'; readonly node: JSONContent }
  | { readonly kind: 'continue'; readonly node: JSONContent }

const listFormats = ['decimal', 'lowerLetter', 'lowerRoman']

function paragraphItems(paragraph: Element, context: Context): Item[] {
  const pPr = kid(paragraph, 'w:pPr')
  if (kid(pPr, 'w:sectPr') !== undefined) context.sections += 1
  const styleId = val(pPr, 'w:pStyle')
  const style = context.styles.paragraph(styleId)
  const direct = readParagraphProps(pPr)
  const props: ParagraphProps = { ...style.paragraph, ...direct }
  const kind = blockKind(styleId, props, context)
  // Code and rule paragraphs get their shading and line from Sumi's own styles.
  if (direct.decorated === true || (style.paragraph.decorated === true && kind !== 'code' && kind !== 'rule')) {
    context.findings.add('Paragraph borders and shading', 'degraded', { alternative: 'Shown and saved without them.' })
  }
  const themeStyle: StyleName = kind === 'title' || kind === 'heading1' || kind === 'heading2' || kind === 'heading3' || kind === 'quote' ? kind : 'normal'
  const base = context.theme[themeStyle]
  const segments = inlineContent(paragraph, base, style.run, context)
  if (props.tabStops === true && segments.some(segment => segment.some(node => node.text?.includes('\t') === true))) {
    context.findings.add('Tab stops', 'degraded', { alternative: 'Tabs line up at the default positions.' })
  }
  const levels = props.numId === undefined || props.numId === '0' ? undefined : context.numbering.get(props.numId)
  const listLevel = levels?.get(props.level ?? 0)
  const isList = listLevel !== undefined && (kind === 'normal' || kind === 'continue')
  if (listLevel !== undefined && !isList && kind.startsWith('heading')) {
    context.findings.add('Numbered headings', 'degraded', { alternative: 'Shown and saved without their numbers.' })
  }
  const fontSize = style.run.size ?? base.fontSize
  const attrs = paragraphAttrs(props, base, fontSize, !isList && kind !== 'checklist' && kind !== 'code' && kind !== 'continue' && kind !== 'cell', context)

  const items: Item[] = []
  if (props.pageBreakBefore === true) items.push({ kind: 'block', node: { type: 'pageBreak' } })
  segments.forEach((content, index) => {
    if (index > 0) items.push({ kind: 'block', node: { type: 'pageBreak' } })
    // A page break at the start or end of a paragraph leaves an empty part that Word never showed.
    if (segments.length > 1 && content.length === 0) return
    const node = (type: string, extra: Record<string, string | number> = {}): JSONContent => {
      const allAttrs = { ...extra, ...attrs }
      return { type, ...(Object.keys(allAttrs).length > 0 ? { attrs: allAttrs } : {}), ...(content.length > 0 ? { content } : {}) }
    }
    if (isList && listLevel !== undefined) {
      const level = props.level ?? 0
      if (listLevel.format !== (listLevel.ordered ? listFormats[level % 3] : 'bullet')) {
        context.findings.add('List numbering styles', 'degraded', { alternative: 'Shown and saved as 1., a., i. and bullets.' })
      }
      items.push({ kind: 'list', numId: props.numId ?? '', level, ordered: listLevel.ordered, start: listLevel.start, node: node('paragraph') })
      return
    }
    switch (kind) {
      case 'title':
        items.push({ kind: 'block', node: node('title') })
        return
      case 'heading1':
      case 'heading2':
      case 'heading3':
        items.push({ kind: 'block', node: node('heading', { level: Number(kind.slice(-1)) }) })
        return
      case 'quote':
        items.push({ kind: 'quote', node: node('paragraph') })
        return
      case 'code':
        items.push({ kind: 'code', text: content.map(child => child.type === 'hardBreak' ? '\n' : child.text ?? '').join('') })
        return
      case 'rule':
        items.push({ kind: 'block', node: { type: 'horizontalRule' } })
        return
      case 'checklist': {
        const [first, ...rest] = content
        const text = first?.text ?? ''
        const checked = text.startsWith(checklistMarks.done)
        const stripped = first !== undefined && /^[☐☑] ?/.test(text)
          ? [...(text.replace(/^[☐☑] ?/, '') === '' ? [] : [{ ...first, text: text.replace(/^[☐☑] ?/, '') }]), ...rest]
          : content
        items.push({ kind: 'checklist', checked, node: { type: 'paragraph', ...(stripped.length > 0 ? { content: stripped } : {}), ...(Object.keys(attrs).length > 0 ? { attrs } : {}) } })
        return
      }
      case 'continue':
        items.push({ kind: 'continue', node: node('paragraph') })
        return
      default:
        items.push({ kind: 'block', node: node('paragraph') })
    }
  })
  return items
}

interface CellBuild {
  readonly header: boolean
  readonly colspan: number
  rowspan: number
  readonly colwidth: number[] | null
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
  if (tableBorders(element)) context.findings.add('Table borders and styles', 'degraded', { alternative: 'Shown and saved with a thin border around every cell.' })
  // Cells merged down (vMerge) grow the cell that starts the merge, by column.
  const merging = new Map<number, CellBuild>()
  const rows = kids(element, 'w:tr').map(row => {
    const header = toggle(kid(row, 'w:trPr'), 'w:tblHeader') === true
    let column = number(val(kid(row, 'w:trPr'), 'w:gridBefore')) ?? 0
    const cells: CellBuild[] = []
    for (const cell of kids(row, 'w:tc')) {
      const tcPr = kid(cell, 'w:tcPr')
      const span = number(val(tcPr, 'w:gridSpan')) ?? 1
      const vMerge = kid(tcPr, 'w:vMerge')
      const origin = merging.get(column)
      if (vMerge !== undefined && attr(vMerge, 'w:val') !== 'restart' && origin !== undefined) {
        origin.rowspan += 1
        column += span
        continue
      }
      const widths = grid.slice(column, column + span)
      const fill = attr(kid(tcPr, 'w:shd'), 'w:fill')
      const content = blocks(cell, context)
      const build: CellBuild = {
        header,
        colspan: span,
        rowspan: 1,
        colwidth: widths.length === span && widths.every(width => width > 0) ? widths : null,
        background: fill === undefined || fill === 'auto' ? undefined : `#${fill.toLowerCase()}`,
        content: content.length > 0 ? content : [{ type: 'paragraph' }],
      }
      if (vMerge !== undefined) merging.set(column, build)
      else merging.delete(column)
      cells.push(build)
      column += span
    }
    return cells
  })
  return {
    type: 'table',
    content: rows.filter(cells => cells.length > 0).map(cells => ({
      type: 'tableRow',
      content: cells.map(cell => ({
        type: cell.header ? 'tableHeader' : 'tableCell',
        attrs: { colspan: cell.colspan, rowspan: cell.rowspan, colwidth: cell.colwidth, ...(cell.background === undefined ? {} : { backgroundColor: cell.background }) },
        content: cell.content,
      })),
    })),
  }
}

/** Groups consecutive list, checklist, quote and code paragraphs into their containers. */
function group(items: readonly Item[], context: Context): JSONContent[] {
  const result: JSONContent[] = []
  let open: Array<{ readonly list: JSONContent; readonly level: number; readonly ordered: boolean }> = []

  /** Counts the item in its list and level, clearing deeper levels as Word does; resolves the count before it. */
  const count = (item: Extract<Item, { kind: 'list' }>): number => {
    const counts = context.listCounts.get(item.numId) ?? []
    const before = counts[item.level] ?? 0
    counts[item.level] = before + 1
    counts.length = item.level + 1
    context.listCounts.set(item.numId, counts)
    return before
  }

  items.forEach((item, index) => {
    const previous = result[result.length - 1]
    if (item.kind !== 'list' && item.kind !== 'continue') open = []
    switch (item.kind) {
      case 'block':
        result.push(item.node)
        return
      case 'code': {
        if (previous?.type === 'codeBlock' && items[index - 1]?.kind === 'code') {
          const text = `${previous.content?.[0]?.text ?? ''}\n${item.text}`
          previous.content = [{ type: 'text', text }]
        } else result.push({ type: 'codeBlock', ...(item.text === '' ? {} : { content: [{ type: 'text', text: item.text }] }) })
        return
      }
      case 'quote':
        if (previous?.type === 'blockquote' && items[index - 1]?.kind === 'quote') previous.content?.push(item.node)
        else result.push({ type: 'blockquote', content: [item.node] })
        return
      case 'checklist': {
        const taskItem = { type: 'taskItem', attrs: { checked: item.checked }, content: [item.node] }
        if (previous?.type === 'taskList' && items[index - 1]?.kind === 'checklist') previous.content?.push(taskItem)
        else result.push({ type: 'taskList', content: [taskItem] })
        return
      }
      case 'continue': {
        const deepest = open[open.length - 1]?.list.content
        const lastItem = deepest?.[deepest.length - 1]
        if (lastItem !== undefined) lastItem.content?.push(item.node)
        else result.push(item.node)
        return
      }
      case 'list': {
        const before = count(item)
        while (open.length > 0 && (open[open.length - 1]?.level ?? 0) > item.level) open.pop()
        const top = open[open.length - 1]
        if (top === undefined || top.level < item.level || top.ordered !== item.ordered) {
          if (top !== undefined && top.level === item.level) open.pop()
          const start = item.start + before
          const list: JSONContent = item.ordered
            ? { type: 'orderedList', ...(start !== 1 ? { attrs: { start } } : {}), content: [] }
            : { type: 'bulletList', content: [] }
          const parent = open[open.length - 1]?.list.content
          const parentItem = parent?.[parent.length - 1]
          if (parentItem !== undefined) parentItem.content?.push(list)
          else result.push(list)
          open.push({ list, level: item.level, ordered: item.ordered })
        }
        open[open.length - 1]?.list.content?.push({ type: 'listItem', content: [item.node] })
        return
      }
    }
  })
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
          items.push({ kind: 'block', node: table(child, context) })
          break
        case 'w:sdt':
          collect(kid(child, 'w:sdtContent') ?? child)
          break
        case 'w:customXml':
          collect(child)
          break
        case 'm:oMathPara':
          context.findings.add('Equations', 'dropped', { unit: 'equation', alternative: removed })
          break
        case 'w:altChunk':
          context.findings.add('Embedded documents', 'dropped', { alternative: removed })
          break
      }
    }
  }
  collect(container)
  return group(items, context)
}

// ─── Theme, page setup, package ─────────────────────────────────────────

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

/** Text in a header or footer outside its fields, and whether it holds a PAGE field. */
function readFooterPart(xml: Element | undefined): { readonly text: boolean; readonly pageNumber: boolean } {
  if (xml === undefined) return { text: false, pageNumber: false }
  const instructions = [...all(xml, 'w:instrText').map(element => element.textContent ?? ''), ...all(xml, 'w:fldSimple').map(element => attr(element, 'w:instr') ?? '')]
  const pageNumber = instructions.some(instruction => fieldName(instruction) === 'PAGE')
  let depth = 0
  let inResult = false
  let text = all(xml, 'w:drawing').length > 0 || all(xml, 'w:pict').length > 0
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
      } else if (child.tagName === 'w:t' && depth === 0 && !inResult && (child.textContent ?? '').trim() !== '') text = true
    }
  }
  return { text, pageNumber }
}

function readPage(body: Element | undefined, context: Context, xmlPart: (name: string) => Element | undefined): PageSetup {
  const section = kid(body, 'w:sectPr')
  if (section === undefined) return defaultPage
  const size = kid(section, 'w:pgSz')
  const margins = kid(section, 'w:pgMar')
  const length = (element: Element | undefined, name: string, fallback: number): number => Math.abs(number(attr(element, name)) ?? fallback)
  if ((number(attr(kid(section, 'w:cols'), 'w:num')) ?? 1) > 1) {
    context.findings.add('Multiple columns', 'dropped', { alternative: 'Shown and saved in one column.' })
  }
  let pageNumbers = false
  let text = false
  for (const reference of [...kids(section, 'w:headerReference'), ...kids(section, 'w:footerReference')]) {
    const part = context.relationships.get(attr(reference, 'r:id') ?? '')
    const read = readFooterPart(part === undefined ? undefined : xmlPart(part.target))
    if (reference.tagName === 'w:footerReference' && read.pageNumber) pageNumbers = true
    if (read.text || (reference.tagName === 'w:headerReference' && read.pageNumber)) text = true
  }
  if (text) {
    context.findings.add('Header and footer text', 'dropped', {
      alternative: pageNumbers ? 'Removed when saved; page numbers are kept. The original file keeps the text.' : removed,
    })
  }
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

export function readDocx(bytes: Uint8Array): DocxRead {
  const parts = unzipSync(bytes)
  const xmlPart = (name: string): Element | undefined => {
    const part = parts[name]
    return part === undefined ? undefined : parseXml(strFromU8(part))
  }
  const document = xmlPart('word/document.xml')
  const body = kid(document, 'w:body')
  if (body === undefined) throw new Error('This file is not a Word document, or it is damaged.')

  const themeFonts = readThemeFonts(xmlPart('word/theme/theme1.xml'))
  const styles = new StyleSheet(xmlPart('word/styles.xml'), themeFonts)
  const findings = new Findings()
  const context: Context = {
    styles,
    numbering: readNumbering(xmlPart('word/numbering.xml')),
    relationships: readRelationships(xmlPart('word/_rels/document.xml.rels')),
    parts,
    themeFonts,
    theme: readTheme(styles),
    findings,
    fields: [],
    listCounts: new Map(),
    sections: 1,
  }

  const comments = xmlPart('word/comments.xml')
  const commentCount = comments === undefined ? 0 : kids(comments, 'w:comment').length
  if (commentCount > 0) findings.add('Comments', 'dropped', { unit: 'comment', count: commentCount, alternative: removed })

  const content = blocks(body, context)
  const page = readPage(body, context, xmlPart)
  if (context.sections > 1) {
    findings.add('Section breaks', 'dropped', { unit: 'section', count: context.sections, alternative: 'Saved as one section, with the last section’s page size and margins.' })
  }
  return {
    content: { type: 'doc', attrs: { page, theme: context.theme }, content: content.length > 0 ? content : [{ type: 'paragraph' }] },
    findings: findings.list(),
  }
}
