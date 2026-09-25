export type ImportSeverity = 'lossless' | 'degraded' | 'dropped'
export type FindingSeverity = Exclude<ImportSeverity, 'lossless'>

export interface ImportFindingInput {
  readonly construct: string
  readonly severity?: FindingSeverity
  readonly location?: string
  readonly suggestedAlternative?: string
}

export interface ImportFinding extends ImportFindingInput {
  readonly severity: FindingSeverity
}

export interface ImportReport {
  readonly id: string
  readonly sourceName: string
  readonly severity: ImportSeverity
  readonly findings: readonly ImportFinding[]
}

/** Creates a conservative report: an unspecified finding is always degraded. */
export function createImportReport(
  sourceName: string,
  findings: readonly ImportFindingInput[],
): ImportReport {
  const normalized = findings.map(finding => ({
    ...finding,
    severity: finding.severity ?? 'degraded',
  }))
  const severity: ImportSeverity = normalized.some(finding => finding.severity === 'dropped')
    ? 'dropped'
    : normalized.length > 0
      ? 'degraded'
      : 'lossless'

  return {
    id: `${sourceName}:${Date.now()}`,
    sourceName,
    severity,
    findings: normalized,
  }
}
