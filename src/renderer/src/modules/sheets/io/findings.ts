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
    return {
      construct,
      severity: group.severity,
      ...(location === undefined ? {} : { location }),
      ...(group.suggestedAlternative === undefined ? {} : { suggestedAlternative: group.suggestedAlternative }),
    }
  })
}
