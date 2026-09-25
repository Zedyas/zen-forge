/**
 * A document's look, as plain data: the paragraph styles (Normal, Title, Heading 1–3, Quote) and
 * the page setup. Every document carries its own copy (the `doc` node's attributes), read from
 * and written to the styles and section of a .docx file. A template is another set of these.
 * Sizes and spacing are in points, page sizes in twips (1/1440 inch, Word's unit), colours in hex.
 */

import { paperForRegion } from '@shared/print'

export interface BlockStyle {
  readonly fontFamily: string
  readonly fontSize: number
  readonly color: string
  readonly bold: boolean
  readonly italic: boolean
  readonly spaceBefore: number
  readonly spaceAfter: number
  /** Multiple of single line spacing, as Word counts it. */
  readonly lineHeight: number
  readonly indentLeft: number
}

export const styleNames = ['normal', 'title', 'heading1', 'heading2', 'heading3', 'quote'] as const
export type StyleName = (typeof styleNames)[number]

export type DocumentTheme = Readonly<Record<StyleName, BlockStyle>>

export const styleLabels: Readonly<Record<StyleName, string>> = {
  normal: 'Normal text',
  title: 'Title',
  heading1: 'Heading 1',
  heading2: 'Heading 2',
  heading3: 'Heading 3',
  quote: 'Quote',
}

const normal: BlockStyle = {
  fontFamily: 'Arial', fontSize: 11, color: '#000000', bold: false, italic: false,
  spaceBefore: 0, spaceAfter: 8, lineHeight: 1.15, indentLeft: 0,
}

/** Google Docs' defaults: Arial 11 pt, which Mac and Windows both have, so files look the same in Word. */
export const defaultTheme: DocumentTheme = {
  normal,
  title: { ...normal, fontSize: 26, spaceAfter: 6, lineHeight: 1 },
  heading1: { ...normal, fontSize: 20, bold: true, spaceBefore: 18, spaceAfter: 6, lineHeight: 1 },
  heading2: { ...normal, fontSize: 16, bold: true, spaceBefore: 14, spaceAfter: 4, lineHeight: 1 },
  heading3: { ...normal, fontSize: 13, bold: true, spaceBefore: 12, spaceAfter: 4, lineHeight: 1 },
  quote: { ...normal, color: '#555555', italic: true, indentLeft: 18 },
}

/** File properties Word shows under File > Info; kept so saving does not erase them. */
export interface DocumentProperties {
  readonly title: string
  readonly subject: string
  readonly creator: string
  readonly keywords: string
  readonly description: string
}

export const emptyProperties: DocumentProperties = { title: '', subject: '', creator: '', keywords: '', description: '' }

export interface PageSetup {
  readonly width: number
  readonly height: number
  readonly marginTop: number
  readonly marginRight: number
  readonly marginBottom: number
  readonly marginLeft: number
  /** Page numbers centred in the footer, in print and in Word. */
  readonly pageNumbers: boolean
}

export const papers = [
  { name: 'Letter', width: 12_240, height: 15_840 },
  { name: 'A4', width: 11_906, height: 16_838 },
  { name: 'Legal', width: 12_240, height: 20_160 },
] as const

export type PaperName = (typeof papers)[number]['name']

const inch = 1_440

function regionPaper(): PaperName {
  try {
    return paperForRegion(new Intl.Locale(navigator.language).maximize().region ?? '').name
  } catch {
    return 'A4'
  }
}

/** Letter where it is the local standard, A4 elsewhere; lengths shown in inches or centimetres to match. */
export const localPaper: PaperName = regionPaper()
export const lengthUnit: 'in' | 'cm' = localPaper === 'Letter' ? 'in' : 'cm'

export function pageFor(paper: PaperName, landscape = false): PageSetup {
  const size = papers.find(candidate => candidate.name === paper) ?? papers[0]
  return {
    width: landscape ? size.height : size.width,
    height: landscape ? size.width : size.height,
    marginTop: inch, marginRight: inch, marginBottom: inch, marginLeft: inch,
    pageNumbers: false,
  }
}

export const defaultPage: PageSetup = pageFor(localPaper)

export function isLandscape(page: PageSetup): boolean {
  return page.width > page.height
}

/** The named paper the page uses in either orientation, or undefined for another size. */
export function paperOf(page: PageSetup): PaperName | undefined {
  const [short, long] = [Math.min(page.width, page.height), Math.max(page.width, page.height)]
  return papers.find(paper => Math.abs(paper.width - short) <= 20 && Math.abs(paper.height - long) <= 20)?.name
}

/** The width text can use, in CSS pixels (96 per inch). */
export function contentWidthPixels(page: PageSetup): number {
  return (page.width - page.marginLeft - page.marginRight) / 15
}

/** Word's "single" spacing is the font's own line height, about 1.15 times the size for Arial. */
export const singleLineHeight = 1.15

/** The theme as CSS custom properties on the page element; docs.css reads them. */
export function themeVariables(theme: DocumentTheme): ReadonlyArray<readonly [string, string]> {
  return styleNames.flatMap(name => {
    const style = theme[name]
    return [
      [`--sumi-${name}-font`, `"${style.fontFamily}"`],
      [`--sumi-${name}-size`, `${style.fontSize}pt`],
      [`--sumi-${name}-color`, style.color],
      [`--sumi-${name}-weight`, style.bold ? '700' : '400'],
      [`--sumi-${name}-style`, style.italic ? 'italic' : 'normal'],
      [`--sumi-${name}-before`, `${style.spaceBefore}pt`],
      [`--sumi-${name}-after`, `${style.spaceAfter}pt`],
      [`--sumi-${name}-line`, String(style.lineHeight * singleLineHeight)],
      [`--sumi-${name}-indent`, `${style.indentLeft}pt`],
    ] as const
  })
}

export const fontFamilies = ['Arial', 'Helvetica', 'Times New Roman', 'Georgia', 'Garamond', 'Courier New', 'Verdana'] as const
export const fontSizes = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72] as const
export const lineSpacings = [1, 1.15, 1.5, 2] as const

function isRecord(value: unknown): value is Readonly<Record<string, unknown>> {
  return typeof value === 'object' && value !== null
}

function isPageSetup(value: unknown): value is PageSetup {
  return isRecord(value)
    && ['width', 'height', 'marginTop', 'marginRight', 'marginBottom', 'marginLeft'].every(key => typeof value[key] === 'number')
    && typeof value['pageNumbers'] === 'boolean'
}

function isBlockStyle(value: unknown): value is BlockStyle {
  return isRecord(value)
    && typeof value['fontFamily'] === 'string' && typeof value['color'] === 'string'
    && typeof value['bold'] === 'boolean' && typeof value['italic'] === 'boolean'
    && ['fontSize', 'spaceBefore', 'spaceAfter', 'lineHeight', 'indentLeft'].every(key => typeof value[key] === 'number')
}

function isTheme(value: unknown): value is DocumentTheme {
  return isRecord(value) && styleNames.every(name => isBlockStyle(value[name]))
}

/** The page setup from the `doc` node's attributes, or the local default. */
export function documentPage(value: unknown): PageSetup {
  return isPageSetup(value) ? value : defaultPage
}

/** The theme from the `doc` node's attributes, or Sumi's default. */
export function documentTheme(value: unknown): DocumentTheme {
  return isTheme(value) ? value : defaultTheme
}

function isProperties(value: unknown): value is DocumentProperties {
  return isRecord(value) && ['title', 'subject', 'creator', 'keywords', 'description'].every(key => typeof value[key] === 'string')
}

/** The file properties from the `doc` node's attributes. */
export function documentProperties(value: unknown): DocumentProperties {
  return isProperties(value) ? value : emptyProperties
}

/** A length or font size in points, from `11`, `'11pt'` or `'14.67px'` (CSS pixels, 96 per inch). */
export function toPoints(value: unknown): number | undefined {
  if (typeof value === 'number') return Number.isFinite(value) ? value : undefined
  if (typeof value !== 'string') return undefined
  const match = /^(-?[\d.]+)\s*(pt|px)?$/.exec(value.trim())
  if (match === null) return undefined
  const amount = Number(match[1])
  if (!Number.isFinite(amount)) return undefined
  return match[2] === 'px' ? Math.round(amount * 0.75 * 100) / 100 : amount
}
