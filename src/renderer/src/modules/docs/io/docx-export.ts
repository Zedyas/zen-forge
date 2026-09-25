/**
 * Writes a Sumi document (TipTap JSON) as .docx with the `docx` library. Structure uses Word's
 * built-in styles (Heading 1–3, numbered and bulleted lists), so Word's navigation pane and list
 * tools work; the few block types Word has no style for (quote, code, rule, checklist) get named
 * paragraph styles that the importer maps back.
 */

import type { JSONContent } from '@tiptap/core'
import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  Packer,
  PageBreak,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  UnderlineType,
  WidthType,
  type ILevelsOptions,
  type IParagraphOptions,
  type ParagraphChild,
} from 'docx'
import { contentWidthPixels, type PageFormat } from '../page'
import { checklistMarks, docxStyleNames, type DocxStyleId } from './docx-styles'
import { fromDataUrl, imageSize, type PixelSize } from './images'

export interface DocxExport {
  readonly bytes: Uint8Array
  /** Images that could not be embedded: linked from the web, or in a format Word cannot show. */
  readonly skippedImages: number
}


type Block = Paragraph | Table

interface ExportContext {
  readonly page: PageFormat
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
  const number = typeof value === 'string' ? Number(value) : value
  return typeof number === 'number' && Number.isFinite(number) && number > 0 ? number : undefined
}

function hasMark(node: JSONContent, type: string): boolean {
  return node.marks?.some(mark => mark.type === type) ?? false
}

function linkHref(node: JSONContent): string | undefined {
  const link = node.marks?.find(mark => mark.type === 'link')
  const href: unknown = link?.attrs?.['href']
  return typeof href === 'string' && href !== '' ? href : undefined
}

const alignments = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
} as const

function alignmentOf(node: JSONContent): (typeof AlignmentType)[keyof typeof AlignmentType] | undefined {
  const align = textAttr(node, 'textAlign')
  return align === 'left' || align === 'center' || align === 'right' || align === 'justify' ? alignments[align] : undefined
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

function inlineRun(node: JSONContent, context: ExportContext, inLink: boolean): TextRun | ImageRun | undefined {
  if (node.type === 'hardBreak') return new TextRun({ break: 1 })
  if (node.type === 'image') return imageRun(node, context)
  if (node.type !== 'text' || node.text === undefined) return undefined
  const style = hasMark(node, 'code') ? 'InlineCode' : inLink ? 'Hyperlink' : undefined
  return new TextRun({
    text: node.text,
    bold: hasMark(node, 'bold') || undefined,
    italics: hasMark(node, 'italic') || undefined,
    strike: hasMark(node, 'strike') || undefined,
    ...(hasMark(node, 'underline') ? { underline: { type: UnderlineType.SINGLE } } : {}),
    ...(style === undefined ? {} : { style }),
  })
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
    const href = linkHref(node)
    const external = href !== undefined && !href.startsWith('#') ? href : undefined
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

function paragraph(node: JSONContent, context: ExportContext, options: IParagraphOptions = {}): Paragraph {
  const alignment = alignmentOf(node)
  return new Paragraph({
    ...options,
    ...(alignment === undefined ? {} : { alignment }),
    children: runs(node.content ?? [], context),
  })
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
    if (index === 0 && child.type === 'paragraph') return [paragraph(child, context, { numbering })]
    return blocks([child], context, 'ListContinue')
  }))
}

function checklistBlocks(list: JSONContent, context: ExportContext): Block[] {
  return (list.content ?? []).flatMap(item => (item.content ?? []).flatMap((child, index) => {
    if (child.type === 'taskList') return checklistBlocks(child, context)
    if (child.type === 'bulletList' || child.type === 'orderedList') return listBlocks(child, context, 1)
    if (index !== 0 || child.type !== 'paragraph') return blocks([child], context, 'ListContinue')
    const mark = item.attrs?.['checked'] === true ? checklistMarks.done : checklistMarks.open
    const alignment = alignmentOf(child)
    return [new Paragraph({
      style: 'Checklist',
      ...(alignment === undefined ? {} : { alignment }),
      children: [new TextRun(`${mark} `), ...runs(child.content ?? [], context)],
    })]
  }))
}

function table(node: JSONContent, context: ExportContext): Table {
  const rows = node.content ?? []
  const columnCount = Math.max(1, ...rows.map(row => (row.content ?? []).reduce((sum, cell) => sum + (numberAttr(cell, 'colspan') ?? 1), 0)))
  const tableWidth = context.page.width - 2 * context.page.margin
  const columnWidth = Math.floor(tableWidth / columnCount)
  return new Table({
    width: { size: tableWidth, type: WidthType.DXA },
    columnWidths: Array.from({ length: columnCount }, () => columnWidth),
    rows: rows.map(row => new TableRow({
      // Written only when set: Word and mammoth read the element's presence, whatever its value.
      ...((row.content ?? []).length > 0 && (row.content ?? []).every(cell => cell.type === 'tableHeader') ? { tableHeader: true } : {}),
      children: (row.content ?? []).map(cell => {
        const columnSpan = numberAttr(cell, 'colspan') ?? 1
        const rowSpan = numberAttr(cell, 'rowspan') ?? 1
        const children = blocks(cell.content ?? [], context, 'TableText')
        return new TableCell({
          width: { size: columnWidth * columnSpan, type: WidthType.DXA },
          ...(columnSpan > 1 ? { columnSpan } : {}),
          ...(rowSpan > 1 ? { rowSpan } : {}),
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
      case 'heading':
        return [paragraph(node, context, { heading: headingLevels[Math.min(3, numberAttr(node, 'level') ?? 1) - 1] })]
      case 'bulletList':
      case 'orderedList':
        return listBlocks(node, context, 0)
      case 'taskList':
        return checklistBlocks(node, context)
      case 'blockquote':
        return blocks(node.content ?? [], context, 'Quote')
      case 'codeBlock': {
        const text = (node.content ?? []).map(child => child.text ?? '').join('')
        // One paragraph per line; the importer joins them back. A blank line keeps a space so it survives.
        return text.split('\n').map(line => new Paragraph({ style: 'Code', children: [new TextRun(line === '' ? ' ' : line)] }))
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

export async function writeDocx(content: JSONContent, page: PageFormat): Promise<DocxExport> {
  const context: ExportContext = { page, listStarts: new Set(), listInstances: 0, skippedImages: 0 }
  const children = blocks(content.content ?? [], context)
  const document = new Document({
    creator: '',
    lastModifiedBy: '',
    styles: {
      default: {
        document: { run: { font: 'Arial', size: 22 }, paragraph: { spacing: { after: 160, line: 276 } } },
        heading1: { run: { font: 'Arial', size: 40, bold: true, color: '000000' }, paragraph: { spacing: { before: 360, after: 120 }, keepNext: true } },
        heading2: { run: { font: 'Arial', size: 32, bold: true, color: '000000' }, paragraph: { spacing: { before: 280, after: 80 }, keepNext: true } },
        heading3: { run: { font: 'Arial', size: 26, bold: true, color: '000000' }, paragraph: { spacing: { before: 240, after: 80 }, keepNext: true } },
        hyperlink: { run: { color: '1155CC', underline: { type: UnderlineType.SINGLE } } },
        listParagraph: { paragraph: { contextualSpacing: true } },
      },
      paragraphStyles: [
        {
          id: 'Quote', name: docxStyleNames.Quote, basedOn: 'Normal', next: 'Normal',
          run: { color: '555555' },
          paragraph: { indent: { left: 360 }, border: { left: { style: BorderStyle.SINGLE, size: 18, color: 'CCCCCC', space: 12 } } },
        },
        {
          id: 'Code', name: docxStyleNames.Code, basedOn: 'Normal', next: 'Normal',
          run: { font: 'Courier New', size: 20 },
          paragraph: { spacing: { after: 0, line: 240 }, shading: { type: ShadingType.CLEAR, fill: 'F3F3F3', color: 'auto' } },
        },
        {
          id: 'HorizontalLine', name: docxStyleNames.HorizontalLine, basedOn: 'Normal', next: 'Normal',
          paragraph: { border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'BFBFBF', space: 1 } } },
        },
        { id: 'ListContinue', name: docxStyleNames.ListContinue, basedOn: 'Normal', paragraph: { indent: { left: 720 } } },
        { id: 'Checklist', name: docxStyleNames.Checklist, basedOn: 'Normal', paragraph: { contextualSpacing: true } },
        { id: 'TableText', name: docxStyleNames.TableText, basedOn: 'Normal', paragraph: { spacing: { after: 0 } } },
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
          levels: numberingLevels(index => ({ format: numberFormats[index % 3], text: `%${index + 1}.`, start: index === 0 ? start : 1 })),
        })),
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: page.width, height: page.height },
          margin: { top: page.margin, right: page.margin, bottom: page.margin, left: page.margin, header: 720, footer: 720 },
        },
      },
      children: children.length > 0 ? children : [new Paragraph({})],
    }],
  })
  return { bytes: new Uint8Array(await Packer.toArrayBuffer(document)), skippedImages: context.skippedImages }
}
