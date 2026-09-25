/**
 * Writes a Sumi document (TipTap JSON) as .docx with the `docx` library. It writes everything the
 * reader (docx-read.ts) reads, so a file Zendo saves opens again unchanged:
 * - the theme as Word's styles (Normal, Title, Heading 1–3, Quote), so Word's navigation pane,
 *   style gallery and list tools work;
 * - direct formatting (fonts, sizes, colours, highlights, spacing, indents, alignment);
 * - the page setup as the section's page size and margins, and page numbers as a footer field.
 * Blocks Word has no style for (code, rule, checklist) get named styles the reader maps back.
 */

import type { JSONContent } from '@tiptap/core'
import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  Footer,
  HeadingLevel,
  HighlightColor,
  ImageRun,
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
  type ILevelsOptions,
  type IParagraphOptions,
  type IParagraphStyleOptions,
  type IRunOptions,
  type ParagraphChild,
} from 'docx'
import { contentWidthPixels, documentPage, documentTheme, type BlockStyle, type PageSetup } from '../theme'
import { checklistMarks, docxStyleNames, hexColor, highlightColors, type DocxStyleId } from './docx-styles'
import { fromDataUrl, imageSize, type PixelSize } from './images'

export interface DocxExport {
  readonly bytes: Uint8Array
  /** Images that could not be embedded: linked from the web, or in a format Word cannot show. */
  readonly skippedImages: number
}

type Block = Paragraph | Table
type Alignment = (typeof AlignmentType)[keyof typeof AlignmentType]

interface ExportContext {
  readonly page: PageSetup
  /** Start values of the numbered lists, one numbering definition each. */
  readonly listStarts: Set<number>
  /** Each numbered list restarts at its start value, so each gets its own numbering instance. */
  listInstances: number
  skippedImages: number
}

function textAttr(node: JSONContent, name: string): string | undefined {
  const value: unknown = node.attrs?.[name]
  return typeof value === 'string' && value !== '' ? value : undefined
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

/** `14pt`, `18.67px` or a plain number of points. */
function fontPoints(value: unknown): number | undefined {
  if (typeof value === 'number') return value
  if (typeof value !== 'string') return undefined
  const match = /^([\d.]+)(pt|px)?$/.exec(value.trim())
  if (match === null) return undefined
  return match[2] === 'px' ? Number(match[1]) * 0.75 : Number(match[1])
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
  const size = fontPoints(style['fontSize'])
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

function inlineRun(node: JSONContent, context: ExportContext, inLink: boolean): TextRun | ImageRun | undefined {
  if (node.type === 'hardBreak') return new TextRun({ break: 1 })
  if (node.type === 'image') return imageRun(node, context)
  if (node.type !== 'text' || node.text === undefined) return undefined
  // Tabs are their own element in Word; inside w:t they would show as a space.
  const children = node.text.split('\t').flatMap((part, index) => [...(index > 0 ? [new Tab()] : []), ...(part === '' ? [] : [part])])
  return new TextRun({ ...runOptions(node, inLink), children })
}

/** Inline content as runs; neighbouring runs with the same web link share one hyperlink. */
function runs(nodes: readonly JSONContent[], context: ExportContext): ParagraphChild[] {
  const result: ParagraphChild[] = []
  let link: { readonly href: string; readonly children: Array<TextRun | ImageRun> } | undefined
  const flush = (): void => {
    if (link !== undefined) result.push(new ExternalHyperlink({ link: link.href, children: link.children }))
    link = undefined
  }
  for (const node of nodes) {
    // Links inside the document (#anchors) have no bookmark to point at in Word, so they stay plain text.
    const href = mark(node, 'link')?.['href']
    const external = typeof href === 'string' && href !== '' && !href.startsWith('#') ? href : undefined
    const run = inlineRun(node, context, external !== undefined)
    if (run === undefined) continue
    if (external === undefined) {
      flush()
      result.push(run)
      continue
    }
    if (link?.href !== external) {
      flush()
      link = { href: external, children: [] }
    }
    link.children.push(run)
  }
  flush()
  return result
}

/** Direct paragraph formatting: the node's own alignment, spacing and indents, over its style. */
function paragraphOptions(node: JSONContent): IParagraphOptions {
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
    ...(Object.keys(indent).length > 0 ? { indent } : {}),
  }
}

function paragraph(node: JSONContent, context: ExportContext, options: IParagraphOptions = {}): Paragraph {
  return new Paragraph({ ...paragraphOptions(node), ...options, children: runs(node.content ?? [], context) })
}

const headingLevels = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3] as const

function listBlocks(list: JSONContent, context: ExportContext, level: number): Block[] {
  const ordered = list.type === 'orderedList'
  const start = numberAttr(list, 'start') ?? 1
  if (ordered) {
    context.listStarts.add(start)
    context.listInstances += 1
  }
  const numbering = { reference: ordered ? `numbered-${start}` : 'bulleted', level: Math.min(level, 8), instance: ordered ? context.listInstances : 0 }
  return (list.content ?? []).flatMap(item => (item.content ?? []).flatMap((child, index) => {
    if (child.type === 'bulletList' || child.type === 'orderedList') return listBlocks(child, context, level + 1)
    // List paragraphs take their indents from the list level, so only alignment is their own.
    if (index === 0 && child.type === 'paragraph') return [new Paragraph({ numbering, ...alignment(child), children: runs(child.content ?? [], context) })]
    return blocks([child], context, 'ListContinue')
  }))
}

function checklistBlocks(list: JSONContent, context: ExportContext): Block[] {
  return (list.content ?? []).flatMap(item => (item.content ?? []).flatMap((child, index) => {
    if (child.type === 'taskList') return checklistBlocks(child, context)
    if (child.type === 'bulletList' || child.type === 'orderedList') return listBlocks(child, context, 1)
    if (index !== 0 || child.type !== 'paragraph') return blocks([child], context, 'ListContinue')
    const box = item.attrs?.['checked'] === true ? checklistMarks.done : checklistMarks.open
    return [new Paragraph({ style: 'Checklist', ...alignment(child), children: [new TextRun(`${box} `), ...runs(child.content ?? [], context)] })]
  }))
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
        const children = blocks(cell.content ?? [], context, 'TableText')
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
        return [paragraph(node, context, { heading: headingLevels[Math.min(3, Math.max(1, numberAttr(node, 'level') ?? 1)) - 1] })]
      case 'bulletList':
      case 'orderedList':
        return listBlocks(node, context, 0)
      case 'taskList':
        return checklistBlocks(node, context)
      case 'blockquote':
        return blocks(node.content ?? [], context, 'Quote')
      case 'codeBlock': {
        const text = (node.content ?? []).map(child => child.text ?? '').join('')
        // One paragraph per line; the reader joins them back.
        return text.split('\n').map(line => new Paragraph({ style: 'Code', children: line === '' ? [] : [new TextRun(line)] }))
      }
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

function pageNumberFooter(): Footer {
  return new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT] })] })] })
}

export async function writeDocx(content: JSONContent): Promise<DocxExport> {
  const page = documentPage(content.attrs?.['page'])
  const theme = documentTheme(content.attrs?.['theme'])
  const context: ExportContext = { page, listStarts: new Set(), listInstances: 0, skippedImages: 0 }
  const children = blocks(content.content ?? [], context)
  const document = new Document({
    creator: '',
    lastModifiedBy: '',
    styles: {
      default: {
        document: styleProperties(theme.normal),
        title: styleProperties(theme.title),
        heading1: styleProperties(theme.heading1, true),
        heading2: styleProperties(theme.heading2, true),
        heading3: styleProperties(theme.heading3, true),
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
      ],
      characterStyles: [
        { id: 'InlineCode', name: docxStyleNames.InlineCode, basedOn: 'DefaultParagraphFont', run: { font: 'Courier New', size: 20 } },
      ],
    },
    numbering: {
      config: [
        { reference: 'bulleted', levels: numberingLevels(index => ({ format: LevelFormat.BULLET, text: bulletGlyphs[index % 3] })) },
        ...[...context.listStarts].map(start => ({
          reference: `numbered-${start}`,
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
      },
      ...(page.pageNumbers ? { footers: { default: pageNumberFooter() } } : {}),
      children: children.length > 0 ? children : [new Paragraph({})],
    }],
  })
  return { bytes: new Uint8Array(await Packer.toArrayBuffer(document)), skippedImages: context.skippedImages }
}
