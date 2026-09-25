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
