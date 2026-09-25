/**
 * The single registry of suite and application identity. Code uses the functional ids
 * (`sheets`, `pdf`, …); every user-facing name comes from here, so renaming is one edit.
 */

export const suiteName = 'Zendo'

export const applicationIds = ['home', 'sheets', 'pdf', 'docs', 'slides'] as const
export type ApplicationId = (typeof applicationIds)[number]
export type EditorApplicationId = Exclude<ApplicationId, 'home'>

export interface ApplicationDefinition {
  readonly id: EditorApplicationId
  readonly name: string
  /** What the application edits, in plain words, for people who do not know the name yet. */
  readonly kind: string
  readonly available: boolean
  /** Extensions it opens, lowercase. The first is the one new documents save as. */
  readonly opens: readonly string[]
  /** Extensions it may write. A file it opens but cannot write is saved under the first of these. */
  readonly saves: readonly string[]
}

export const applications: readonly ApplicationDefinition[] = [
  { id: 'sheets', name: 'Ledger', kind: 'Spreadsheets', available: true, opens: ['xlsx', 'xlsm', 'csv', 'tsv', 'xls', 'ods', 'numbers'], saves: ['xlsx', 'csv', 'tsv'] },
  { id: 'pdf', name: 'Hanko', kind: 'PDF editor', available: true, opens: ['pdf'], saves: ['pdf'] },
  { id: 'docs', name: 'Sumi', kind: 'Documents', available: true, opens: ['docx', 'md', 'markdown'], saves: ['docx', 'md', 'markdown'] },
  { id: 'slides', name: 'Slides', kind: 'Presentations', available: true, opens: ['pptx'], saves: ['pptx'] },
]

/** Images are read (never written) so they can be combined into a PDF. */
export const imageExtensions: readonly string[] = ['png', 'jpg', 'jpeg']

export function findApplication(id: EditorApplicationId): ApplicationDefinition {
  const application = applications.find(candidate => candidate.id === id)
  if (application === undefined) throw new Error(`Unknown application: ${id}`)
  return application
}

export function applicationForExtension(extension: string): EditorApplicationId | undefined {
  const normalized = extension.toLowerCase()
  return applications.find(application => application.opens.includes(normalized))?.id
}

/** Every extension the desktop shell may read; guards the filesystem IPC surface. */
export function isReadableExtension(extension: string): boolean {
  const normalized = extension.toLowerCase()
  return applicationForExtension(normalized) !== undefined || imageExtensions.includes(normalized)
}

/** Every extension the desktop shell may write. */
export function isWritableExtension(extension: string): boolean {
  const normalized = extension.toLowerCase()
  return applications.some(application => application.saves.includes(normalized))
}
