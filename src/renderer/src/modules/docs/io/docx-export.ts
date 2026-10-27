/**
 * Writes a Sumi document (TipTap JSON) as .docx with the `docx` library, in the form the reader
 * (docx-read.ts) reads back:
 * - the theme as Word's styles (Normal, Title, Heading 1–6, Quote), so Word's navigation pane,
 *   style gallery and list tools work;
 * - formatting on the text as direct formatting (fonts, sizes, colours, highlights, spacing,
 *   indents, alignment), and the file properties;
 * - the page setup as the section's page size and margins, and the headers and footers as the
 *   section's, with page numbers as PAGE and NUMPAGES fields;
 * - a table of contents as Word's TOC field, with today's headings as its result and bookmarks on
 *   the headings for its links, so Word shows it at once and can update it;
 * - blocks Word has no style for (code, rule, checklist) as named styles, and a block that
 *   continues a list item as a list level with no number, so it stays inside the item.
 * Whatever does not come back the same (a table inside a list item, say) is caught when saving
 * over a file: the save reads its own output first and asks (doc-documents.ts).
 */

import type { JSONContent } from '@tiptap/core'
import {
  AlignmentType,
  BookmarkEnd,
  BookmarkStart,
  BorderStyle,
  BuilderElement,
  Document,
  ExternalHyperlink,
  Footer,
  Header,
  HeadingLevel,
  HighlightColor,
  ImageRun,
  InternalHyperlink,
  LevelFormat,
  LineRuleType,
  Packer,
  PageBreak,
  PageNumber,
  Paragraph,
  ShadingType,
  Tab,
  Table,
  TableCell,
  TableRow,
  TextRun,
  UnderlineType,
  WidthType,
  XmlComponent,
  type ILevelsOptions,
  type IParagraphOptions,
  type IParagraphStyleOptions,
  type IRunOptions,
} from 'docx'
import { documentHeadersAndFooters, type HeaderFooterContent } from '../header-footer'
import { contentsLevels } from '../text-stats'
import { contentWidthPixels, documentPage, documentProperties, documentTheme, toPoints, type BlockStyle, type PageSetup } from '../theme'
import { checklistMarks, contentsInstruction, docxStyleNames, hexColor, highlightColors, safeHref, type DocxStyleId } from './docx-styles'
import { fromDataUrl, imageSize, type PixelSize } from './images'

export interface DocxExport {
  readonly bytes: Uint8Array
  /** Images that could not be embedded: linked from the web, or in a format Word cannot show. */
  readonly skippedImages: number
}

type Block = Paragraph | Table
type Inline = TextRun | ImageRun | PageBreak | ExternalHyperlink
type Alignment = (typeof AlignmentType)[keyof typeof AlignmentType]

/** A heading as a table of contents lists it, with the bookmark its entry links to. */
interface ContentsEntry {
  readonly level: number
  readonly text: string
  readonly bookmark: string
}

interface ExportContext {
  readonly page: PageSetup
  /** Every heading in order, when the document has a table of contents; empty otherwise. */
  readonly headings: readonly ContentsEntry[]
  /** Headings written so far, which numbers each one's bookmark. */
  headingCount: number
  /**
   * Each numbered list's own numbering definition and its start value. Word counts on across lists
   * that share a definition, and only a list's first level can restart, so a nested list needs its own.
   */
  readonly numberedLists: Map<string, number>
  /** Each numbered list restarts at its start value, so each gets its own numbering instance. */
  listInstances: number
  skippedImages: number
}

/**
 * Text as XML 1.0 allows it: control characters other than tab and line ends make a file Word
 * cannot open. Callers turn line and page breaks into Word's own breaks first.
 */
export function xmlText(text: string): string {
  // eslint-disable-next-line no-control-regex
  return text.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, '')
}

function textAttr(node: JSONContent, name: string): string | undefined {
  const value: unknown = node.attrs?.[name]
  return typeof value === 'string' && value !== '' ? xmlText(value) : undefined
}

function numberAttr(node: JSONContent, name: string): number | undefined {
  const value: unknown = node.attrs?.[name]
  const parsed = typeof value === 'string' ? Number(value) : value
  return typeof parsed === 'number' && Number.isFinite(parsed) ? parsed : undefined
}

function mark(node: JSONContent, type: string): Readonly<Record<string, unknown>> | undefined {
  const found = node.marks?.find(candidate => candidate.type === type)
  return found === undefined ? undefined : found.attrs ?? {}
}

function twips(points: number): number {
  return Math.round(points * 20)
}

/** A colour as Word writes it: six hex digits, no `#`. */
function wordColor(value: unknown): string | undefined {
  return typeof value === 'string' ? hexColor(value)?.slice(1).toUpperCase() : undefined
}

const alignments: Readonly<Record<string, Alignment>> = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
}

function alignment(node: JSONContent): IParagraphOptions {
  const align = alignments[textAttr(node, 'textAlign') ?? '']
  return align === undefined ? {} : { alignment: align }
}

/**
 * The size an image shows at: its own width, or the width it was given, scaled down to fit the
 * text column. The editor's CSS (max-width: 100%) does the same.
 */
export function displaySize(node: JSONContent, natural: PixelSize, columnWidth: number): PixelSize {
  const width = numberAttr(node, 'width') ?? natural.width
  const height = numberAttr(node, 'height') ?? (natural.width > 0 ? natural.height * (width / natural.width) : natural.height)
  const scale = width > columnWidth ? columnWidth / width : 1
  return { width: Math.round(width * scale), height: Math.round(height * scale) }
}

const imageTypes: Readonly<Record<string, 'png' | 'jpg' | 'gif' | 'bmp'>> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
}

function imageRun(node: JSONContent, context: ExportContext): ImageRun | undefined {
  const decoded = fromDataUrl(textAttr(node, 'src') ?? '')
  const type = decoded === undefined ? undefined : imageTypes[decoded.mimeType]
  const natural = decoded === undefined ? undefined : imageSize(decoded.bytes)
  if (decoded === undefined || type === undefined || natural === undefined) {
    context.skippedImages += 1
    return undefined
  }
  const alt = textAttr(node, 'alt')
  return new ImageRun({
    type,
    data: decoded.bytes,
    transformation: displaySize(node, natural, contentWidthPixels(context.page)),
    ...(alt === undefined ? {} : { altText: { name: alt, description: alt, title: alt } }),
  })
}

type Highlight = (typeof HighlightColor)[keyof typeof HighlightColor]

/** Word's highlight name for a colour that is one of its highlight colours. */
const highlightNames = new Map<string, Highlight>(Object.values(HighlightColor).flatMap(name => {
  const color = highlightColors[name]
  return color === undefined ? [] : [[color, name] as const]
}))

/** Character formatting from a text node's marks. Styles supply the rest. */
function runOptions(node: JSONContent, inLink: boolean): IRunOptions {
  if (mark(node, 'code') !== undefined) return { style: 'InlineCode' }
  const style = mark(node, 'textStyle') ?? {}
  const size = toPoints(style['fontSize'])
  const font = style['fontFamily']
  const color = wordColor(style['color'])
  const background = wordColor(style['backgroundColor'])
  const highlight = background === undefined ? undefined : highlightNames.get(`#${background.toLowerCase()}`)
  return {
    ...(mark(node, 'bold') === undefined ? {} : { bold: true }),
    ...(mark(node, 'italic') === undefined ? {} : { italics: true }),
    ...(mark(node, 'strike') === undefined ? {} : { strike: true }),
    ...(mark(node, 'underline') === undefined ? {} : { underline: { type: UnderlineType.SINGLE } }),
    ...(mark(node, 'superscript') === undefined ? {} : { superScript: true }),
    ...(mark(node, 'subscript') === undefined ? {} : { subScript: true }),
    ...(typeof font === 'string' && font !== '' ? { font } : {}),
    ...(size === undefined ? {} : { size: Math.round(size * 2) }),
    ...(color === undefined ? {} : { color }),
    ...(highlight !== undefined ? { highlight } : background !== undefined ? { shading: { type: ShadingType.CLEAR, fill: background, color: 'auto' } } : {}),
    ...(inLink ? { style: 'Hyperlink' } : {}),
  }
}

/**
 * A text node as runs. Tabs are their own element in Word (inside w:t they show as a space); a
 * vertical tab or line end pasted into the text becomes a line break, a form feed a page break.
 */
function textRuns(node: JSONContent, text: string, inLink: boolean): Array<TextRun | PageBreak> {
  return text.split('\f').flatMap((page, pageIndex) => {
    const lines = page.split(/\r\n|[\v\n\r]/).map((line, lineIndex) => {
      const children = xmlText(line).split('\t').flatMap((part, index) => [...(index > 0 ? [new Tab()] : []), ...(part === '' ? [] : [part])])
      return new TextRun({ ...runOptions(node, inLink), ...(lineIndex > 0 ? { break: 1 } : {}), children })
    })
    return pageIndex > 0 ? [new PageBreak(), ...lines] : lines
  })
}

function inlineRuns(node: JSONContent, context: ExportContext, inLink: boolean): Array<TextRun | ImageRun | PageBreak> {
  if (node.type === 'hardBreak') return [new TextRun({ break: 1 })]
  if (node.type === 'pageField') return [new TextRun({ ...runOptions(node, inLink), children: [node.attrs?.['kind'] === 'total' ? PageNumber.TOTAL_PAGES : PageNumber.CURRENT] })]
  if (node.type === 'image') {
    const image = imageRun(node, context)
    return image === undefined ? [] : [image]
  }
  if (node.type !== 'text' || node.text === undefined) return []
  return textRuns(node, node.text, inLink)
}

/** Inline content as runs; neighbouring runs with the same web link share one hyperlink. */
function runs(nodes: readonly JSONContent[], context: ExportContext): Inline[] {
  const result: Inline[] = []
  let link: { readonly href: string; readonly children: Array<TextRun | ImageRun | PageBreak> } | undefined
  const flush = (): void => {
    if (link !== undefined) result.push(new ExternalHyperlink({ link: link.href, children: link.children }))
    link = undefined
  }
  for (const node of nodes) {
    // Only web and mail links are written; links inside the document (#anchors) have no bookmark
    // to point at in Word, and other schemes are never kept, so those stay plain text.
    const href = mark(node, 'link')?.['href']
    const external = typeof href === 'string' ? safeHref(xmlText(href)) : undefined
    const inline = inlineRuns(node, context, external !== undefined)
    if (external === undefined) {
      flush()
      result.push(...inline)
      continue
    }
    if (link?.href !== external) {
      flush()
      link = { href: external, children: [] }
    }
    link.children.push(...inline)
  }
  flush()
  return result
}

/**
 * Direct paragraph formatting: the node's own alignment, spacing and indents, over its style.
 * A paragraph in a list takes its indent from the list level, so `indents` is off there.
 */
function paragraphOptions(node: JSONContent, indents = true): IParagraphOptions {
  const before = numberAttr(node, 'spaceBefore')
  const after = numberAttr(node, 'spaceAfter')
  const line = numberAttr(node, 'lineHeight')
  const left = numberAttr(node, 'indentLeft')
  const firstLine = numberAttr(node, 'indentFirstLine')
  const spacing = {
    ...(before === undefined ? {} : { before: twips(before) }),
    ...(after === undefined ? {} : { after: twips(after) }),
    ...(line === undefined ? {} : { line: Math.round(line * 240), lineRule: LineRuleType.AUTO }),
  }
  const indent = {
    ...(left === undefined ? {} : { left: twips(left) }),
    ...(firstLine === undefined ? {} : firstLine < 0 ? { hanging: twips(-firstLine) } : { firstLine: twips(firstLine) }),
  }
  return {
    ...alignment(node),
    ...(Object.keys(spacing).length > 0 ? { spacing } : {}),
    ...(indents && Object.keys(indent).length > 0 ? { indent } : {}),
  }
}

function paragraph(node: JSONContent, context: ExportContext, options: IParagraphOptions = {}): Paragraph {
  return new Paragraph({ ...paragraphOptions(node), ...options, children: runs(node.content ?? [], context) })
}

/** A heading, inside a bookmark when a table of contents links to it. */
function heading(node: JSONContent, context: ExportContext): Paragraph {
  const level = Math.min(6, Math.max(1, numberAttr(node, 'level') ?? 1))
  const entry = context.headings[context.headingCount]
  context.headingCount += 1
  const children = runs(node.content ?? [], context)
  const paragraph = new Paragraph({ ...paragraphOptions(node), heading: headingLevels[level - 1], children: entry === undefined ? children : [] })
  if (entry !== undefined) {
    // The docx library's Bookmark numbers every bookmark 1; Word needs each one's id to differ.
    paragraph.addChildElement(new BookmarkStart(entry.bookmark, context.headingCount))
    children.forEach(child => paragraph.addChildElement(child))
    paragraph.addChildElement(new BookmarkEnd(context.headingCount))
  }
  return paragraph
}

/** The headings in document order, the order `blocks` writes them in. */
function contentsEntries(nodes: readonly JSONContent[]): ContentsEntry[] {
  const entries: ContentsEntry[] = []
  const visit = (node: JSONContent): void => {
    if (node.type === 'heading') {
      const text = xmlText((node.content ?? []).map(child => child.text ?? '').join('').replace(/[\t\n\r\f\v]/g, ' '))
      entries.push({ level: numberAttr(node, 'level') ?? 1, text, bookmark: `_Toc${entries.length + 1}` })
    }
    node.content?.forEach(visit)
  }
  nodes.forEach(visit)
  return entries
}

/** Part of a field (`w:fldChar` or `w:instrText`); the docx library writes these only inside its own fields. */
function fieldCharacter(type: 'begin' | 'separate' | 'end'): XmlComponent {
  return new BuilderElement<{ readonly type: string }>({ name: 'w:fldChar', attributes: { type: { key: 'w:fldCharType', value: type } } })
}

class FieldCode extends XmlComponent {
  constructor(code: string) {
    super('w:instrText')
    this.root.push(code)
  }
}

/**
 * A table of contents: the field starts in the first entry's paragraph and ends in the last's,
 * and its result is the entries as they are now, each linking to its heading's bookmark. Page
 * numbers are left for Word to add when the field is updated.
 */
function tableOfContents(node: JSONContent, context: ExportContext): Paragraph[] {
  const levels = contentsLevels(node.attrs?.['levels'])
  const begin = new TextRun({ children: [fieldCharacter('begin'), new FieldCode(contentsInstruction(levels)), fieldCharacter('separate')] })
  const end = new TextRun({ children: [fieldCharacter('end')] })
  const entries = context.headings.filter(entry => entry.level <= levels)
  if (entries.length === 0) return [new Paragraph({ children: [begin, end] })]
  return entries.map((entry, index) => new Paragraph({
    style: `TOC${entry.level}`,
    children: [
      ...(index === 0 ? [begin] : []),
      new InternalHyperlink({ anchor: entry.bookmark, children: [new TextRun(entry.text)] }),
      ...(index === entries.length - 1 ? [end] : []),
    ],
  }))
}

const headingLevels = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6] as const

type Numbering = NonNullable<IParagraphOptions['numbering']>

/** A list level with no number: marks a block that continues the list item it follows, at that level. */
function placement(level: number): Numbering {
  return { reference: 'placement', level: Math.min(level, 8), instance: 0 }
}

/** A paragraph, code block or other block inside a list item, after its first paragraph. */
function continuation(node: JSONContent, context: ExportContext, level: number, style: DocxStyleId | undefined): Block[] {
  const numbering = placement(level)
  if (node.type === 'paragraph') {
    return [new Paragraph({ ...paragraphOptions(node, false), style: style ?? 'ListContinue', numbering, children: runs(node.content ?? [], context) })]
  }
  if (node.type === 'codeBlock') return codeLines(node).map(line => new Paragraph({ style: 'Code', numbering, children: line === '' ? [] : [new TextRun(line)] }))
  if (node.type === 'bulletList' || node.type === 'orderedList') return listBlocks(node, context, level + 1, style)
  if (node.type === 'taskList') return checklistBlocks(node, context, level + 1)
  return blocks([node], context, style)
}

/** Bulleted and numbered lists; each list is its own Word list, so neighbouring lists stay apart. */
function listBlocks(list: JSONContent, context: ExportContext, level: number, style?: DocxStyleId): Block[] {
  const ordered = list.type === 'orderedList'
  context.listInstances += 1
  const reference = ordered ? `numbered-${context.listInstances}` : 'bulleted'
  if (ordered) context.numberedLists.set(reference, numberAttr(list, 'start') ?? 1)
  const numbering = { reference, level: Math.min(level, 8), instance: context.listInstances }
  return (list.content ?? []).flatMap(item => (item.content ?? []).flatMap((child, index) => {
    if (index === 0 && child.type === 'paragraph') {
      return [new Paragraph({ ...paragraphOptions(child, false), numbering, ...(style === undefined ? {} : { style }), children: runs(child.content ?? [], context) })]
    }
    return continuation(child, context, level, style)
  }))
}

/** Word has no checklist; an item is a paragraph that starts with a box, at its level of nesting. */
function checklistBlocks(list: JSONContent, context: ExportContext, level: number): Block[] {
  const numbering = { reference: 'checklist', level: Math.min(level, 8), instance: 0 }
  return (list.content ?? []).flatMap(item => (item.content ?? []).flatMap((child, index) => {
    if (index !== 0 || child.type !== 'paragraph') return continuation(child, context, level, undefined)
    const box = item.attrs?.['checked'] === true ? checklistMarks.done : checklistMarks.open
    return [new Paragraph({ ...paragraphOptions(child, false), style: 'Checklist', numbering, children: [new TextRun(`${box} `), ...runs(child.content ?? [], context)] })]
  }))
}

function codeLines(node: JSONContent): string[] {
  return xmlText((node.content ?? []).map(child => child.text ?? '').join('').replace(/\r\n|[\v\r\f]/g, '\n')).split('\n')
}

/** Column widths in twips: the widths the first row's cells were given, the rest sharing what is left of the text width. */
function columnWidths(rows: readonly JSONContent[], page: PageSetup): number[] {
  const given = (rows[0]?.content ?? []).flatMap(cell => {
    const widths: unknown = cell.attrs?.['colwidth']
    const span = numberAttr(cell, 'colspan') ?? 1
    return Array.isArray(widths) && widths.length === span && widths.every(width => typeof width === 'number' && width > 0)
      ? widths.map(width => Number(width) * 15)
      : Array.from({ length: span }, () => 0)
  })
  const count = Math.max(1, ...rows.map(row => (row.content ?? []).reduce((sum, cell) => sum + (numberAttr(cell, 'colspan') ?? 1), 0)))
  const available = page.width - page.marginLeft - page.marginRight
  const known = given.filter(width => width > 0)
  const share = count > known.length ? Math.max(360, (available - known.reduce((sum, width) => sum + width, 0)) / (count - known.length)) : 0
  return Array.from({ length: count }, (_, index) => Math.round((given[index] ?? 0) > 0 ? given[index] ?? 0 : share))
}

function table(node: JSONContent, context: ExportContext): Table {
  const rows = node.content ?? []
  const widths = columnWidths(rows, context.page)
  return new Table({
    width: { size: widths.reduce((sum, width) => sum + width, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map(row => new TableRow({
      // Written only when set: Word and readers go by the element's presence, whatever its value.
      ...((row.content ?? []).length > 0 && (row.content ?? []).every(cell => cell.type === 'tableHeader') ? { tableHeader: true } : {}),
      children: (row.content ?? []).map(cell => {
        const columnSpan = numberAttr(cell, 'colspan') ?? 1
        const rowSpan = numberAttr(cell, 'rowspan') ?? 1
        const shading = wordColor(cell.attrs?.['backgroundColor'])
        // A column's alignment (from Markdown) is the alignment of each paragraph in its cells.
        const align = textAttr(cell, 'align')
        const content = (cell.content ?? []).map(child => align !== undefined && child.type === 'paragraph' && textAttr(child, 'textAlign') === undefined
          ? { ...child, attrs: { ...child.attrs, textAlign: align } }
          : child)
        const children = blocks(content, context, 'TableText')
        return new TableCell({
          ...(columnSpan > 1 ? { columnSpan } : {}),
          ...(rowSpan > 1 ? { rowSpan } : {}),
          ...(shading === undefined ? {} : { shading: { type: ShadingType.CLEAR, fill: shading, color: 'auto' } }),
          // Word requires a paragraph in every cell.
          children: children.length > 0 ? children : [new Paragraph({ style: 'TableText' })],
        })
      }),
    })),
  })
}

/** Converts block nodes; `style` is the paragraph style of the container (quote, table cell, list). */
function blocks(nodes: readonly JSONContent[], context: ExportContext, style?: DocxStyleId): Block[] {
  return nodes.flatMap((node): Block[] => {
    switch (node.type) {
      case 'paragraph':
        return [paragraph(node, context, style === undefined ? {} : { style })]
      case 'title':
        return [paragraph(node, context, { heading: HeadingLevel.TITLE })]
      case 'heading':
        return [heading(node, context)]
      case 'tableOfContents':
        return tableOfContents(node, context)
      case 'bulletList':
      case 'orderedList':
        return listBlocks(node, context, 0, style === 'Quote' ? style : undefined)
      case 'taskList':
        return checklistBlocks(node, context, 0)
      case 'blockquote':
        return blocks(node.content ?? [], context, 'Quote')
      case 'codeBlock':
        // One paragraph per line; the reader joins them back.
        return codeLines(node).map(line => new Paragraph({ style: 'Code', children: line === '' ? [] : [new TextRun(line)] }))
      case 'horizontalRule':
        return [new Paragraph({ style: 'HorizontalLine' })]
      case 'pageBreak':
        return [new Paragraph({ children: [new PageBreak()] })]
      case 'table':
        return [table(node, context)]
      case 'image':
        return [new Paragraph({ children: runs([node], context) })]
      default:
        return []
    }
  })
}

function numberingLevels(level: (index: number) => Omit<ILevelsOptions, 'level' | 'alignment' | 'style'>): ILevelsOptions[] {
  return Array.from({ length: 9 }, (_, index) => ({
    level: index,
    alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 720 * (index + 1), hanging: 360 } } },
    ...level(index),
  }))
}

/** Levels that only indent: for blocks continuing a list item, and for checklists. */
function unnumberedLevels(indent: (index: number) => number): ILevelsOptions[] {
  return Array.from({ length: 9 }, (_, index) => ({
    level: index,
    format: LevelFormat.NONE,
    text: '',
    alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: indent(index), hanging: 0 } } },
  }))
}

const bulletGlyphs = ['•', '◦', '▪'] as const
const numberFormats = [LevelFormat.DECIMAL, LevelFormat.LOWER_LETTER, LevelFormat.LOWER_ROMAN] as const

/** A theme style as Word style properties. */
function styleProperties(style: BlockStyle, keepNext = false): Required<Pick<IParagraphStyleOptions, 'run' | 'paragraph'>> {
  return {
    run: {
      font: style.fontFamily,
      size: Math.round(style.fontSize * 2),
      color: style.color.slice(1).toUpperCase(),
      ...(style.bold ? { bold: true } : {}),
      ...(style.italic ? { italics: true } : {}),
    },
    paragraph: {
      spacing: { before: twips(style.spaceBefore), after: twips(style.spaceAfter), line: Math.round(style.lineHeight * 240), lineRule: LineRuleType.AUTO },
      ...(style.indentLeft === 0 ? {} : { indent: { left: twips(style.indentLeft) } }),
      ...(keepNext ? { keepNext: true } : {}),
    },
  }
}

/** A header or footer's paragraphs, in Word's Header or Footer style. */
function partParagraphs(content: HeaderFooterContent, style: 'Header' | 'Footer', context: ExportContext): Paragraph[] {
  return content.map(node => new Paragraph({ ...alignment(node), style, children: runs(node.content ?? [], context) }))
}

export async function writeDocx(content: JSONContent): Promise<DocxExport> {
  const page = documentPage(content.attrs?.['page'])
  const theme = documentTheme(content.attrs?.['theme'])
  const properties = documentProperties(content.attrs?.['properties'])
  const hasContents = (nodes: readonly JSONContent[]): boolean => nodes.some(node => node.type === 'tableOfContents' || hasContents(node.content ?? []))
  const context: ExportContext = {
    page,
    headings: hasContents(content.content ?? []) ? contentsEntries(content.content ?? []) : [],
    headingCount: 0,
    numberedLists: new Map(),
    listInstances: 0,
    skippedImages: 0,
  }
  const children = blocks(content.content ?? [], context)
  const headerFooter = documentHeadersAndFooters(content.attrs?.['headerFooter'])
  // The first page's own header and footer are written even while unused, as Word keeps them.
  const header = (part: HeaderFooterContent): Header | undefined => part.length === 0 ? undefined : new Header({ children: partParagraphs(part, 'Header', context) })
  const footer = (part: HeaderFooterContent): Footer | undefined => part.length === 0 ? undefined : new Footer({ children: partParagraphs(part, 'Footer', context) })
  const document = new Document({
    title: xmlText(properties.title),
    subject: xmlText(properties.subject),
    creator: xmlText(properties.creator),
    keywords: xmlText(properties.keywords),
    description: xmlText(properties.description),
    lastModifiedBy: '',
    styles: {
      default: {
        document: styleProperties(theme.normal),
        title: styleProperties(theme.title),
        heading1: styleProperties(theme.heading1, true),
        heading2: styleProperties(theme.heading2, true),
        heading3: styleProperties(theme.heading3, true),
        heading4: styleProperties(theme.heading4, true),
        heading5: styleProperties(theme.heading5, true),
        heading6: styleProperties(theme.heading6, true),
        hyperlink: { run: { color: '1155CC', underline: { type: UnderlineType.SINGLE } } },
        listParagraph: { paragraph: { contextualSpacing: true } },
      },
      paragraphStyles: [
        { id: 'Quote', name: docxStyleNames.Quote, basedOn: 'Normal', next: 'Normal', ...styleProperties(theme.quote) },
        {
          id: 'Code', name: docxStyleNames.Code, basedOn: 'Normal', next: 'Normal',
          run: { font: 'Courier New', size: 20 },
          paragraph: { spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO }, shading: { type: ShadingType.CLEAR, fill: 'F3F3F3', color: 'auto' } },
        },
        {
          id: 'HorizontalLine', name: docxStyleNames.HorizontalLine, basedOn: 'Normal', next: 'Normal',
          paragraph: { border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'BFBFBF', space: 1 } } },
        },
        { id: 'ListContinue', name: docxStyleNames.ListContinue, basedOn: 'Normal', paragraph: { indent: { left: 720 } } },
        { id: 'Checklist', name: docxStyleNames.Checklist, basedOn: 'Normal', paragraph: { contextualSpacing: true } },
        { id: 'TableText', name: docxStyleNames.TableText, basedOn: 'Normal', paragraph: { spacing: { before: 0, after: 0 } } },
        { id: 'Header', name: 'header', basedOn: 'Normal', paragraph: { spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO } } },
        { id: 'Footer', name: 'footer', basedOn: 'Normal', paragraph: { spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO } } },
        // Word's table of contents styles: each level indented a little more than the one above.
        ...[1, 2, 3, 4, 5, 6].map(level => ({ id: `TOC${level}`, name: `toc ${level}`, basedOn: 'Normal', next: 'Normal', paragraph: { indent: { left: 220 * (level - 1) }, spacing: { after: 100 } } })),
      ],
      characterStyles: [
        { id: 'InlineCode', name: docxStyleNames.InlineCode, basedOn: 'DefaultParagraphFont', run: { font: 'Courier New', size: 20 } },
      ],
    },
    numbering: {
      config: [
        { reference: 'bulleted', levels: numberingLevels(index => ({ format: LevelFormat.BULLET, text: bulletGlyphs[index % 3] })) },
        { reference: 'placement', levels: unnumberedLevels(index => 720 * (index + 1)) },
        { reference: 'checklist', levels: unnumberedLevels(index => 360 * index) },
        ...[...context.numberedLists].map(([reference, start]) => ({
          reference,
          levels: numberingLevels(index => ({ format: numberFormats[index % 3], text: `%${index + 1}.`, start })),
        })),
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: page.width, height: page.height },
          margin: { top: page.marginTop, right: page.marginRight, bottom: page.marginBottom, left: page.marginLeft, header: 720, footer: 720 },
        },
        ...(headerFooter.differentFirstPage ? { titlePage: true } : {}),
      },
      headers: { default: header(headerFooter.header), first: header(headerFooter.firstHeader) },
      footers: { default: footer(headerFooter.footer), first: footer(headerFooter.firstFooter) },
      children: children.length > 0 ? children : [new Paragraph({})],
    }],
  })
  return { bytes: new Uint8Array(await Packer.toArrayBuffer(document)), skippedImages: context.skippedImages }
}
