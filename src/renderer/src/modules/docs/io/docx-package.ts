/**
 * mammoth converts what it understands and silently skips the rest, so the importer also reads
 * the raw package: what it finds here is what the next save would lose or change.
 */

import { strFromU8, unzipSync } from 'fflate'
import type { ImportFindingInput } from '@shared/fidelity'
import type { PageFormat } from '../page'
import type { PixelSize } from './images'

interface DocxParts {
  readonly names: readonly string[]
  readonly document: string
  readonly styles: string
  readonly numbering: string
  readonly theme: string
  readonly comments: string
  /** Every header and footer part, concatenated. */
  readonly headersAndFooters: string
}

export interface DocxInspection {
  readonly findings: readonly ImportFindingInput[]
  /** The size each picture is placed at, in document order, for the importer to apply. */
  readonly imageSizes: readonly PixelSize[]
}

const headerOrFooter = /^word\/(header|footer)\d*\.xml$/

function readParts(bytes: Uint8Array): DocxParts {
  const names: string[] = []
  const wanted = new Set(['word/document.xml', 'word/styles.xml', 'word/numbering.xml', 'word/theme/theme1.xml', 'word/comments.xml'])
  const inflated = unzipSync(bytes, {
    filter: file => {
      if (!file.name.endsWith('/')) names.push(file.name)
      return wanted.has(file.name) || headerOrFooter.test(file.name)
    },
  })
  const text = (name: string): string => {
    const part = inflated[name]
    return part === undefined ? '' : strFromU8(part)
  }
  if (inflated['word/document.xml'] === undefined) throw new Error('This file is not a Word document, or it is damaged.')
  return {
    names,
    document: text('word/document.xml'),
    styles: text('word/styles.xml'),
    numbering: text('word/numbering.xml'),
    theme: text('word/theme/theme1.xml'),
    comments: text('word/comments.xml'),
    headersAndFooters: names.filter(name => headerOrFooter.test(name)).map(text).join(''),
  }
}

function attribute(tag: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1]
}

/** The body text's defaults: Word's document defaults, then the Normal style. */
function bodyDefaults(styles: string): string {
  const defaults = /<w:docDefaults>[\s\S]*?<\/w:docDefaults>/.exec(styles)?.[0] ?? ''
  const normal = /<w:style\b[^>]*w:styleId="Normal"[^>]*>[\s\S]*?<\/w:style>/.exec(styles)?.[0] ?? ''
  return defaults + normal
}

/** Fonts other than Arial that the body text or directly formatted text uses. */
function otherFonts(parts: DocxParts): string[] {
  const themeFont = (kind: 'minor' | 'major'): string | undefined =>
    new RegExp(`<a:${kind}Font>[\\s\\S]*?<a:latin typeface="([^"]*)"`).exec(parts.theme)?.[1]
  const fonts = new Set<string>()
  for (const [tag] of (parts.document + bodyDefaults(parts.styles)).matchAll(/<w:rFonts\b[^>]*>/g)) {
    for (const name of ['w:ascii', 'w:hAnsi']) {
      const font = attribute(tag, name)
      if (font !== undefined) fonts.add(font)
    }
    for (const name of ['w:asciiTheme', 'w:hAnsiTheme']) {
      const theme = attribute(tag, name)
      const font = theme === undefined ? undefined : themeFont(theme.startsWith('major') ? 'major' : 'minor')
      if (font !== undefined) fonts.add(font)
    }
  }
  fonts.delete('Arial')
  fonts.delete('')
  return [...fonts]
}

function fontSizes(parts: DocxParts): boolean {
  const sizes = [...(parts.document + bodyDefaults(parts.styles)).matchAll(/<w:sz w:val="(\d+)"/g)].map(match => match[1])
  return sizes.some(size => size !== '22')
}

/** Field codes other than hyperlinks (which mammoth keeps as links): TOC, PAGE, REF, DATE… */
function fieldNames(document: string): string[] {
  const instructions = [
    ...[...document.matchAll(/<w:instrText\b[^>]*>([^<]*)<\/w:instrText>/g)].map(match => match[1] ?? ''),
    ...[...document.matchAll(/<w:fldSimple\b[^>]*w:instr="([^"]*)"/g)].map(match => match[1] ?? ''),
  ]
  const names = new Set(instructions.map(instruction => instruction.trim().split(/\s+/)[0]?.toUpperCase() ?? '').filter(name => name !== ''))
  if (/<w:docPartGallery w:val="Table of Contents"/.test(document)) names.add('TOC')
  names.delete('HYPERLINK')
  return [...names]
}

function pageDiffers(document: string, page: PageFormat): boolean {
  // The body's own section properties are the last ones in the document.
  const section = [...document.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].at(-1)?.[0]
  if (section === undefined) return false
  const size = /<w:pgSz\b[^>]*>/.exec(section)?.[0] ?? ''
  const margins = /<w:pgMar\b[^>]*>/.exec(section)?.[0] ?? ''
  const near = (value: string | undefined, expected: number): boolean => value === undefined || Math.abs(Number(value) - expected) <= 20
  return !near(attribute(size, 'w:w'), page.width) || !near(attribute(size, 'w:h'), page.height)
    || ['w:top', 'w:right', 'w:bottom', 'w:left'].some(side => !near(attribute(margins, side), page.margin))
}

/** Where each picture is placed, from its drawing extent (EMU) or VML style (points). */
function pictureSizes(document: string): PixelSize[] {
  const sizes: PixelSize[] = []
  for (const [, kind, body = ''] of document.matchAll(/<w:(drawing|pict)>([\s\S]*?)<\/w:\1>/g)) {
    if (kind === 'drawing') {
      if (!body.includes('<a:blip')) continue
      const extent = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(body)
      if (extent !== null) sizes.push({ width: Math.round(Number(extent[1]) / 9525), height: Math.round(Number(extent[2]) / 9525) })
    } else if (body.includes('<v:imagedata')) {
      const style = /style="([^"]*)"/.exec(body)?.[1] ?? ''
      const width = /width:([\d.]+)pt/.exec(style)?.[1]
      const height = /height:([\d.]+)pt/.exec(style)?.[1]
      if (width !== undefined && height !== undefined) {
        sizes.push({ width: Math.round(Number(width) * 4 / 3), height: Math.round(Number(height) * 4 / 3) })
      }
    }
  }
  return sizes
}

interface Detector {
  readonly construct: string
  readonly severity: 'degraded' | 'dropped'
  readonly suggestedAlternative?: string
  readonly found: (parts: DocxParts) => boolean
}

const detectors: readonly Detector[] = [
  {
    construct: 'Tracked changes',
    severity: 'dropped',
    suggestedAlternative: 'Shown with every change accepted',
    found: parts => /<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange)\b/.test(parts.document),
  },
  {
    construct: 'Comments',
    severity: 'dropped',
    found: parts => /<w:comment\b/.test(parts.comments) || /<w:commentReference\b/.test(parts.document),
  },
  {
    construct: 'Footnotes and endnotes',
    severity: 'dropped',
    suggestedAlternative: 'Their text is kept as a numbered list at the end',
    found: parts => /<w:(footnoteReference|endnoteReference)\b/.test(parts.document),
  },
  {
    construct: 'Headers and footers',
    severity: 'dropped',
    found: parts => /<w:t(\s[^>]*)?>[^<]|<w:drawing\b|<w:pict\b|<w:fldSimple\b|<w:instrText\b/.test(parts.headersAndFooters),
  },
  {
    construct: 'Text boxes and shapes',
    severity: 'dropped',
    suggestedAlternative: 'Text in text boxes is kept as paragraphs',
    found: parts => /<w:txbxContent\b|<wps:wsp\b|<wpg:wgp\b|<v:(rect|oval|line|roundrect|polyline|arc|textbox)\b/.test(parts.document),
  },
  {
    construct: 'Embedded objects',
    severity: 'dropped',
    found: parts => /<w:object\b|<o:OLEObject\b/.test(parts.document) || parts.names.some(name => name.startsWith('word/embeddings/')),
  },
  { construct: 'Charts', severity: 'dropped', found: parts => parts.names.some(name => name.startsWith('word/charts/')) },
  { construct: 'SmartArt', severity: 'dropped', found: parts => parts.names.some(name => name.startsWith('word/diagrams/')) },
  { construct: 'Equations', severity: 'dropped', found: parts => /<m:oMath\b/.test(parts.document) },
  {
    construct: 'Form fields and check boxes',
    severity: 'dropped',
    found: parts => /<w:ffData\b|<w14:checkbox\b/.test(parts.document),
  },
  {
    construct: 'Multiple columns',
    severity: 'dropped',
    suggestedAlternative: 'Shown in one column',
    found: parts => [...parts.document.matchAll(/<w:cols\b[^>]*w:num="(\d+)"/g)].some(match => Number(match[1]) > 1),
  },
  {
    construct: 'Images in EMF, WMF or TIFF format',
    severity: 'dropped',
    found: parts => parts.names.some(name => /^word\/media\/.+\.(emf|wmf|tiff?)$/i.test(name)),
  },
  {
    construct: 'Font sizes',
    severity: 'degraded',
    suggestedAlternative: 'Body text is shown at 11 pt',
    found: fontSizes,
  },
  {
    construct: 'Text colours',
    severity: 'degraded',
    found: parts => /<w:color w:val="(?!auto"|000000")[^"]*"/i.test(parts.document),
  },
  {
    construct: 'Highlights and shading',
    severity: 'degraded',
    found: parts => /<w:highlight\b|<w:shd\b[^>]*w:fill="(?!auto"|FFFFFF")[0-9A-F]{6}"/i.test(parts.document),
  },
  {
    construct: 'Paragraph spacing and indents',
    severity: 'degraded',
    found: parts => /<w:(spacing|ind)\b/.test(parts.document),
  },
  {
    construct: 'Tab stops',
    severity: 'degraded',
    suggestedAlternative: 'Tabs are shown as spaces',
    found: parts => /<w:tab\/>/.test(parts.document),
  },
  {
    construct: 'Section breaks',
    severity: 'degraded',
    suggestedAlternative: 'Shown as one continuous section',
    found: parts => (parts.document.match(/<w:sectPr\b/g) ?? []).length > 1,
  },
  {
    construct: 'List numbering styles',
    severity: 'degraded',
    suggestedAlternative: 'Shown as 1, a, i and bullets',
    found: parts => /<w:numFmt w:val="(?!decimal"|bullet"|lowerLetter"|lowerRoman"|none")/.test(parts.numbering),
  },
  {
    construct: 'Title and subtitle styles',
    severity: 'degraded',
    suggestedAlternative: 'Shown as Heading 1 and normal text',
    found: parts => /<w:pStyle w:val="(Title|Subtitle)"/.test(parts.document),
  },
  {
    construct: 'Image positions and text wrapping',
    severity: 'degraded',
    suggestedAlternative: 'Shown in line with the text',
    found: parts => /<wp:anchor\b/.test(parts.document),
  },
]

/** Lists what the package holds that Sumi does not keep, plus the placed size of each picture. */
export function inspectDocx(bytes: Uint8Array, page: PageFormat): DocxInspection {
  const parts = readParts(bytes)
  const findings: ImportFindingInput[] = detectors
    .filter(detector => detector.found(parts))
    .map(({ construct, severity, suggestedAlternative }) => ({ construct, severity, ...(suggestedAlternative === undefined ? {} : { suggestedAlternative }) }))

  const fields = fieldNames(parts.document)
  if (fields.length > 0) {
    findings.push({ construct: 'Fields such as a table of contents', severity: 'dropped', location: fields.join(', '), suggestedAlternative: 'Their current text is kept' })
  }
  const fonts = otherFonts(parts)
  if (fonts.length > 0) findings.push({ construct: 'Fonts', severity: 'degraded', location: fonts.join(', '), suggestedAlternative: 'Shown in Arial' })
  if (pageDiffers(parts.document, page)) {
    findings.push({ construct: 'Page size and margins', severity: 'degraded', suggestedAlternative: `Shown on ${page.name} paper with 1-inch margins` })
  }
  return { findings, imageSizes: pictureSizes(parts.document) }
}
