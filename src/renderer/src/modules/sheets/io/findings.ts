/** Import findings are raised per occurrence and collapsed into one finding per construct. */

import type { FindingSeverity, ImportFindingInput } from '@shared/fidelity'

export interface FindingEvent {
  readonly construct: string
  readonly severity: FindingSeverity
  /** Sheet name, defined name or package part the construct was found in. */
  readonly location?: string
  readonly count?: number
  /** Plural noun for `count`, e.g. `cells`. Without it the location is reported on its own. */
  readonly unit?: string
  readonly suggestedAlternative?: string
}

/**
 * What happens to each construct, in the words the import report shows under its name. An event's
 * own `suggestedAlternative` wins; every construct an importer raises should have an entry here.
 */
const consequences: Readonly<Record<string, string>> = {
  'Macros (VBA)': 'Never run here, and not kept in a file Zendo saves.',
  'Pivot tables': 'Their last results show as ordinary cells; the pivot table itself is not kept when saved.',
  'Charts': 'Not shown, and not kept when saved.',
  'Images and drawings': 'Not shown, and not kept when saved.',
  'External links': 'Cells keep their last values; the links to other files are not kept.',
  'Comments': 'Not shown, and not kept when saved.',
  'Form controls': 'Not shown, and not kept when saved.',
  'Slicers': 'Not shown, and not kept when saved.',
  'External data connections': 'Cells keep their last values; the connection is not kept, so the data no longer refreshes.',
  'Unrecognized workbook content': 'Not kept when saved.',
  'Rich text formatting': 'Each cell shows its text in one style; mixed formatting inside a cell is not kept.',
  'Hyperlinks': 'The link text is kept; the links are not.',
  'Array formulas': 'Calculated as ordinary formulas, which can give different results.',
  'Strikethrough, outline and script fonts': 'Shown as regular text, and saved that way.',
  'Font family and size': 'Shown in the default font and size, and saved that way.',
  'Theme and indexed colours': 'Shown in the default colour, and saved that way.',
  'Gradient fills': 'Shown without the fill, and saved that way.',
  'Pattern fills': 'Shown without the fill, and saved that way.',
  'Other text alignment': 'Justified, indented and rotated text shows and saves with the default alignment.',
  'Text wrapping': 'Long text stays on one line, and saves without wrapping.',
  'Vertical alignment': 'Shown and saved with the default vertical alignment.',
  'Unmapped number formats': 'Shown in a simpler format here. The original format is saved unchanged unless you change the cell\'s number format.',
  'Cell borders': 'Not shown, and not kept when saved.',
  'Custom row heights': 'Rows show at the standard height, and save that way.',
  'Split panes': 'Shown without the split.',
  'Merged cells': 'Shown as separate cells, and saved unmerged. Their values are kept.',
  'Conditional formatting': 'Not shown, and not kept when saved.',
  'Data validation': 'Rules and dropdown lists are not checked here, and not kept when saved.',
  'Tables': 'The data is kept; the table\'s name, style and filter buttons are not.',
  'Sheet protection': 'The sheet can be edited here, and saves without protection.',
  'Auto filters': 'Filter buttons are not shown, and not kept when saved.',
  'Hidden sheets': 'Shown as ordinary sheets, and saved visible.',
  'Print areas and titles': 'Not kept when saved.',
  'Defined names that are not a single range': 'Not kept when saved; formulas that use them may show errors.',
  'Charts and embedded objects': 'Not shown, and not kept when saved.',
  'Chart sheets': 'Not shown, and not kept when saved.',
}

interface ConstructGroup {
  severity: FindingSeverity
  unit?: string
  suggestedAlternative?: string
  readonly locations: Map<string, number>
}

function describe(locations: Map<string, number>, unit: string | undefined): string | undefined {
  if (locations.size === 0) return undefined
  return [...locations]
    .map(([location, count]) => {
      if (unit === undefined) return location
      return `${location} (${count} ${count === 1 ? unit.replace(/s$/, '') : unit})`
    })
    .join(', ')
}

/** One finding per construct, with each location counted, so a file never reports per-cell noise. */
export function aggregateFindings(events: readonly FindingEvent[]): ImportFindingInput[] {
  const groups = new Map<string, ConstructGroup>()

  for (const event of events) {
    const group: ConstructGroup = groups.get(event.construct)
      ?? { severity: event.severity, locations: new Map<string, number>() }
    if (event.severity === 'dropped') group.severity = 'dropped'
    group.unit ??= event.unit
    group.suggestedAlternative ??= event.suggestedAlternative
    if (event.location !== undefined) {
      group.locations.set(event.location, (group.locations.get(event.location) ?? 0) + (event.count ?? 1))
    }
    groups.set(event.construct, group)
  }

  return [...groups].map(([construct, group]) => {
    const location = describe(group.locations, group.unit)
    const explanation = group.suggestedAlternative ?? consequences[construct]
    return {
      construct,
      severity: group.severity,
      ...(location === undefined ? {} : { location }),
      ...(explanation === undefined ? {} : { suggestedAlternative: explanation }),
    }
  })
}
