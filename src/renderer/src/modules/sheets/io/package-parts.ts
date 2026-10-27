/**
 * ExcelJS silently discards whole parts of an xlsx package, so the importer inspects the raw zip
 * as well: what it finds here is what would disappear on the next save.
 */

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import type { FindingEvent } from './findings'

/** Font 0 of `xl/styles.xml`, which every cell inherits unless its own style overrides it. */
export interface DefaultFont {
  readonly name?: string
  readonly size?: number
  readonly colorTheme?: number
}

export interface WorkbookPackage {
  readonly parts: readonly string[]
  readonly defaultFont: DefaultFont
}

/** Notes and the VML shapes that place them, which ExcelJS reads as cell notes. */
const notesPart = /^xl\/(comments\d*\.xml|drawings\/(_rels\/)?vmlDrawing\d*\.vml(\.rels)?)$/

interface PartDetector {
  readonly construct: string
  readonly severity: 'degraded' | 'dropped'
  readonly matches: (part: string) => boolean
}

const detectors: readonly PartDetector[] = [
  { construct: 'Macros (VBA)', severity: 'dropped', matches: part => part === 'xl/vbaProject.bin' },
  {
    construct: 'Pivot tables',
    severity: 'dropped',
    matches: part => part.startsWith('xl/pivotTables/') || part.startsWith('xl/pivotCache/'),
  },
  { construct: 'Charts', severity: 'dropped', matches: part => part.startsWith('xl/charts/') },
  {
    construct: 'Images and drawings',
    severity: 'dropped',
    matches: part => (part.startsWith('xl/drawings/') && !notesPart.test(part)) || part.startsWith('xl/media/'),
  },
  { construct: 'External links', severity: 'dropped', matches: part => part.startsWith('xl/externalLinks/') },
  {
    construct: 'Threaded comments',
    severity: 'dropped',
    matches: part => part.startsWith('xl/threadedComments/') || part.startsWith('xl/persons/'),
  },
  {
    construct: 'Form controls',
    severity: 'dropped',
    matches: part => part.startsWith('xl/ctrlProps/') || part.startsWith('xl/activeX/'),
  },
  {
    construct: 'Slicers',
    severity: 'dropped',
    matches: part => part.startsWith('xl/slicers/') || part.startsWith('xl/slicerCaches/'),
  },
  {
    construct: 'External data connections',
    severity: 'dropped',
    matches: part => part === 'xl/connections.xml' || part.startsWith('xl/queryTables/'),
  },
]

/** Parts the importer reads through ExcelJS, or that carry nothing the model could hold. */
const readParts = new Set([
  'xl/workbook.xml',
  'xl/styles.xml',
  'xl/sharedStrings.xml',
  'xl/calcChain.xml',
  'xl/metadata.xml',
])
const readPrefixes = ['xl/_rels/', 'xl/theme/', 'xl/worksheets/', 'xl/tables/', 'xl/printerSettings/']

function parseDefaultFont(xml: string): DefaultFont {
  const fonts = xml.match(/<fonts[^>]*>([\s\S]*?)<\/fonts>/)?.[1]
  const first = fonts?.match(/<font[^>]*>([\s\S]*?)<\/font>/)?.[1]
  if (first === undefined) return {}

  const size = first.match(/<sz val="([\d.]+)"/)?.[1]
  const theme = first.match(/<color[^>]*\btheme="(\d+)"/)?.[1]
  return {
    name: first.match(/<name val="([^"]*)"/)?.[1],
    size: size === undefined ? undefined : Number(size),
    colorTheme: theme === undefined ? undefined : Number(theme),
  }
}

/** Lists the package entries without inflating them, and inflates only the style sheet. */
export function readWorkbookPackage(bytes: Uint8Array): WorkbookPackage {
  const parts: string[] = []
  const inflated = unzipSync(bytes, {
    filter: file => {
      if (!file.name.endsWith('/')) parts.push(file.name)
      return file.name === 'xl/styles.xml'
    },
  })

  const styles = inflated['xl/styles.xml']
  return { parts, defaultFont: styles === undefined ? {} : parseDefaultFont(strFromU8(styles)) }
}

function attribute(tag: string, name: string): string | undefined {
  return tag.match(new RegExp(`\\s${name}="([^"]*)"`))?.[1]
}

const entities: Readonly<Record<string, string>> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }
const escapes: Readonly<Record<string, string>> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }

function unescapeXml(text: string): string {
  return text.replace(/&(?:#x([\da-f]+)|#(\d+)|(\w+));/gi, (match, hex?: string, decimal?: string, name?: string) =>
    hex !== undefined ? String.fromCodePoint(parseInt(hex, 16))
      : decimal !== undefined ? String.fromCodePoint(Number(decimal))
        : entities[name ?? ''] ?? match)
}

function escapeXml(text: string): string {
  return text.replace(/[&<>"]/g, character => escapes[character] ?? character)
}

/** Relationship targets by id, as a relationships part lists them (still XML-escaped). */
function relationshipTargets(xml: Uint8Array | undefined): ReadonlyMap<string | undefined, string | undefined> {
  return new Map(xml === undefined ? [] : [...strFromU8(xml).matchAll(/<Relationship\b[^>]*>/g)]
    .map(([tag]) => [attribute(tag, 'Id'), attribute(tag, 'Target')]))
}

/** The relationships part of a part: `xl/worksheets/_rels/sheet1.xml.rels` for `xl/worksheets/sheet1.xml`. */
function relationshipsPart(part: string): string {
  return part.replace(/[^/]+$/, name => `_rels/${name}.rels`)
}

/** Each worksheet's part, by sheet name, as the workbook and its relationships list them. */
function worksheetParts(files: Readonly<Record<string, Uint8Array>>): ReadonlyMap<string, string> {
  const workbook = files['xl/workbook.xml']
  const targets = relationshipTargets(files['xl/_rels/workbook.xml.rels'])
  const parts = new Map<string, string>()
  for (const [tag] of workbook === undefined ? [] : strFromU8(workbook).matchAll(/<sheet\b[^>]*>/g)) {
    const name = attribute(tag, 'name')
    const target = targets.get(attribute(tag, 'r:id'))
    if (name === undefined || target === undefined) continue
    const path = unescapeXml(target)
    parts.set(unescapeXml(name), path.startsWith('/') ? path.slice(1) : `xl/${path}`)
  }
  return parts
}

/**
 * Zero-based rows whose height the file sets on purpose (`customHeight`), by sheet name. ExcelJS reads
 * a row's height but not this mark, which tells a chosen height from one Excel sized by itself.
 */
export function customHeightRows(bytes: Uint8Array): ReadonlyMap<string, ReadonlySet<number>> {
  const files = unzipSync(bytes, {
    filter: file => file.name === 'xl/workbook.xml' || file.name === 'xl/_rels/workbook.xml.rels' || /^xl\/worksheets\/[^/]+\.xml$/.test(file.name),
  })
  const result = new Map<string, ReadonlySet<number>>()
  for (const [name, part] of worksheetParts(files)) {
    const xml = files[part]
    if (xml === undefined) continue
    const rows = [...strFromU8(xml).matchAll(/<row\b[^>]*>/g)]
      .flatMap(([row]) => /\scustomHeight="(?:1|true)"/.test(row) ? [Number(attribute(row, 'r')) - 1] : [])
    result.set(name, new Set(rows))
  }
  return result
}

const hyperlinkType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const emptyRelationships = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
/** The worksheet parts that come after its links, in the order the file format requires. */
const afterHyperlinks = /<(?:printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b|<\/worksheet>/

/**
 * Each sheet's links, by sheet name, then cell address (`B2`). A link to a place in the workbook (a cell,
 * range or name), which the file keeps in a `location` rather than a relationship, starts with `#`.
 */
export type SheetLinks = ReadonlyMap<string, ReadonlyMap<string, string>>

/**
 * Each sheet's links, and the package without them for ExcelJS to read: it would turn every linked
 * cell into text (a linked formula loses its formula) and drop links on cells the file has no element for.
 */
export function takeHyperlinks(bytes: Uint8Array): { readonly bytes: Uint8Array; readonly links: SheetLinks } {
  const files = unzipSync(bytes)
  const links = new Map<string, ReadonlyMap<string, string>>()
  for (const [name, part] of worksheetParts(files)) {
    const xml = files[part] === undefined ? '' : strFromU8(files[part])
    const element = xml.match(/<hyperlinks>[\s\S]*?<\/hyperlinks>/)?.[0]
    if (element === undefined) continue
    const targets = relationshipTargets(files[relationshipsPart(part)])
    const sheetLinks = new Map<string, string>()
    for (const [tag] of element.matchAll(/<hyperlink\b[^>]*>/g)) {
      const cell = attribute(tag, 'ref')?.split(':')[0]
      const location = attribute(tag, 'location')
      const target = targets.get(attribute(tag, 'r:id')) ?? (location === undefined || location === '' ? undefined : `#${location}`)
      if (cell !== undefined && target !== undefined && target !== '') sheetLinks.set(cell, unescapeXml(target))
    }
    links.set(name, sheetLinks)
    files[part] = strToU8(xml.replace(element, ''))
  }
  // Stored, not compressed: ExcelJS reads it once, straight away.
  return { bytes: links.size === 0 ? bytes : zipSync(files, { level: 0 }), links }
}

/** The package ExcelJS wrote, with each sheet's links added: ExcelJS can link only text cells. */
export function withHyperlinks(bytes: Uint8Array, links: SheetLinks): Uint8Array {
  if (![...links.values()].some(sheetLinks => sheetLinks.size > 0)) return bytes
  const files = unzipSync(bytes)
  for (const [name, part] of worksheetParts(files)) {
    const sheetLinks = [...links.get(name) ?? []]
    const xml = files[part] === undefined ? '' : strFromU8(files[part])
    const at = xml.search(afterHyperlinks)
    if (sheetLinks.length === 0 || at < 0) continue
    const elements: string[] = []
    const relationships: string[] = []
    sheetLinks.forEach(([cell, target], index) => {
      if (target.startsWith('#')) {
        elements.push(`<hyperlink ref="${cell}" location="${escapeXml(target.slice(1))}"/>`)
      } else {
        elements.push(`<hyperlink ref="${cell}" r:id="rIdLink${index}"/>`)
        relationships.push(`<Relationship Id="rIdLink${index}" Type="${hyperlinkType}" Target="${escapeXml(target)}" TargetMode="External"/>`)
      }
    })
    files[part] = strToU8(`${xml.slice(0, at)}<hyperlinks>${elements.join('')}</hyperlinks>${xml.slice(at)}`)
    if (relationships.length === 0) continue
    const relationshipsXml = files[relationshipsPart(part)]
    const existing = relationshipsXml === undefined ? emptyRelationships : strFromU8(relationshipsXml)
    files[relationshipsPart(part)] = strToU8(existing.replace('</Relationships>', `${relationships.join('')}</Relationships>`))
  }
  return zipSync(files)
}

export function packageFindings(parts: readonly string[]): readonly FindingEvent[] {
  const events: FindingEvent[] = []

  for (const detector of detectors) {
    const count = parts.filter(detector.matches).length
    if (count > 0) events.push({ construct: detector.construct, severity: detector.severity })
  }

  for (const part of parts) {
    if (!part.startsWith('xl/')) continue
    if (readParts.has(part) || readPrefixes.some(prefix => part.startsWith(prefix)) || notesPart.test(part)) continue
    if (detectors.some(detector => detector.matches(part))) continue
    events.push({ construct: 'Unrecognized workbook content', severity: 'degraded', location: part })
  }

  return events
}
