/**
 * ExcelJS silently discards whole parts of an xlsx package, so the importer inspects the raw zip
 * as well: what it finds here is what would disappear on the next save.
 */

import { strFromU8, unzipSync } from 'fflate'
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
    matches: part => part.startsWith('xl/drawings/') || part.startsWith('xl/media/'),
  },
  { construct: 'External links', severity: 'dropped', matches: part => part.startsWith('xl/externalLinks/') },
  {
    construct: 'Comments',
    severity: 'dropped',
    matches: part => /^xl\/(comments\d*\.xml|threadedComments\/|persons\/)/.test(part),
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

function unescapeXml(text: string): string {
  return text.replace(/&(?:#x([\da-f]+)|#(\d+)|(\w+));/gi, (match, hex?: string, decimal?: string, name?: string) =>
    hex !== undefined ? String.fromCodePoint(parseInt(hex, 16))
      : decimal !== undefined ? String.fromCodePoint(Number(decimal))
        : entities[name ?? ''] ?? match)
}

/**
 * Zero-based rows whose height the file sets on purpose (`customHeight`), by sheet name. ExcelJS reads
 * a row's height but not this mark, which tells a chosen height from one Excel sized by itself.
 */
export function customHeightRows(bytes: Uint8Array): ReadonlyMap<string, ReadonlySet<number>> {
  const files = unzipSync(bytes, {
    filter: file => file.name === 'xl/workbook.xml' || file.name === 'xl/_rels/workbook.xml.rels' || /^xl\/worksheets\/[^/]+\.xml$/.test(file.name),
  })
  const workbook = files['xl/workbook.xml']
  const relationships = files['xl/_rels/workbook.xml.rels']
  if (workbook === undefined || relationships === undefined) return new Map()
  const targets = new Map([...strFromU8(relationships).matchAll(/<Relationship\b[^>]*>/g)]
    .map(([tag]) => [attribute(tag, 'Id'), attribute(tag, 'Target')]))
  const result = new Map<string, ReadonlySet<number>>()
  for (const [tag] of strFromU8(workbook).matchAll(/<sheet\b[^>]*>/g)) {
    const name = attribute(tag, 'name')
    const target = targets.get(attribute(tag, 'r:id'))
    const xml = target === undefined ? undefined : files[target.startsWith('/') ? target.slice(1) : `xl/${target}`]
    if (name === undefined || xml === undefined) continue
    const rows = [...strFromU8(xml).matchAll(/<row\b[^>]*>/g)]
      .flatMap(([row]) => /\scustomHeight="(?:1|true)"/.test(row) ? [Number(attribute(row, 'r')) - 1] : [])
    result.set(unescapeXml(name), new Set(rows))
  }
  return result
}

export function packageFindings(parts: readonly string[]): readonly FindingEvent[] {
  const events: FindingEvent[] = []

  for (const detector of detectors) {
    const count = parts.filter(detector.matches).length
    if (count > 0) events.push({ construct: detector.construct, severity: detector.severity })
  }

  for (const part of parts) {
    if (!part.startsWith('xl/')) continue
    if (readParts.has(part) || readPrefixes.some(prefix => part.startsWith(prefix))) continue
    if (detectors.some(detector => detector.matches(part))) continue
    events.push({ construct: 'Unrecognized workbook content', severity: 'degraded', location: part })
  }

  return events
}
