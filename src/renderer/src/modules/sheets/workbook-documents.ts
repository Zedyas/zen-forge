import { useEffect, useSyncExternalStore } from 'react'
import { toast } from 'sonner'
import { findApplication, isWritableExtension } from '@shared/applications'
import { createImportReport } from '@shared/fidelity'
import { openFiles } from '../../app/document-actions'
import { useDocumentsStore, type OpenDocument } from '../../app/documents-store'
import { fileService } from '../../services/file/IpcFileService'
import { useFidelityStore } from '../../services/fidelity/fidelity-store'
import { recordPreview, recordRecent } from '../../services/index/document-index'
import { openPrintDialog, printToPdf } from '../../services/print/print'
import { isDelimitedFormat, readDelimited, readXlsx, writeDelimited, writeXlsx } from './io'
import { formatCellValue } from './model/format'
import { Workbook } from './model/Workbook'
import { renderPreview } from './preview'
import { sheetPrintout, sheetPrintTable, type SheetPrintTable } from './print'

export type WorkbookEntry =
  | { readonly status: 'loading' }
  | { readonly status: 'ready'; readonly workbook: Workbook }
  | { readonly status: 'error'; readonly message: string }

interface CachedWorkbook {
  entry: WorkbookEntry
  /** The document id this workbook reports dirtiness to; changes when Save As re-keys the tab. */
  documentId: string
  unsubscribe?: () => void
}

// Workbooks outlive the grid component (tabs remount on switch), so they live here, by document id.
const cache = new Map<string, CachedWorkbook>()
const listeners = new Set<() => void>()
let version = 0

function emit(): void {
  version += 1
  listeners.forEach(listener => listener())
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

function setReady(cached: CachedWorkbook, workbook: Workbook): void {
  cached.entry = { status: 'ready', workbook }
  cached.unsubscribe = workbook.subscribe(() => {
    useDocumentsStore.getState().setDirty(cached.documentId, workbook.isDirty())
  })
  emit()
}

async function load(document: OpenDocument, cached: CachedWorkbook): Promise<void> {
  if (document.path === undefined) {
    setReady(cached, new Workbook())
    return
  }
  try {
    const bytes = await fileService.read(document.path)
    const imported = isDelimitedFormat(document.extension)
      ? readDelimited(bytes, document.name, document.extension)
      : await readXlsx(bytes)
    const findings = document.extension === 'xlsm'
      ? [...imported.findings, { construct: 'Macro-enabled file format (.xlsm)', severity: 'dropped' as const, suggestedAlternative: 'Saved as a new .xlsx file; macros are not kept' }]
      : imported.findings
    useFidelityStore.getState().publish(document.id, createImportReport(`${document.name}.${document.extension}`, findings))
    setReady(cached, new Workbook(imported.data))
  } catch (error) {
    cached.entry = { status: 'error', message: error instanceof Error ? error.message : 'The file could not be read.' }
    emit()
  }
}

/** The workbook for a tab, loading it from disk on first use. */
export function useWorkbookEntry(document: OpenDocument): WorkbookEntry {
  useSyncExternalStore(subscribe, () => version)
  const cached = cache.get(document.id)
  useEffect(() => {
    if (cache.has(document.id)) return
    const created: CachedWorkbook = { entry: { status: 'loading' }, documentId: document.id }
    cache.set(document.id, created)
    void load(document, created)
  }, [document])
  return cached?.entry ?? { status: 'loading' }
}

function workbookFor(documentId: string): Workbook | undefined {
  const entry = cache.get(documentId)?.entry
  return entry?.status === 'ready' ? entry.workbook : undefined
}

export function releaseWorkbook(documentId: string): void {
  cache.get(documentId)?.unsubscribe?.()
  cache.delete(documentId)
  useFidelityStore.getState().forget(documentId)
  emit()
}

/** The visible values of a sheet's used area, exactly as the grid formats them. */
function displayRows(workbook: Workbook, sheetId: number): string[][] {
  const { width, height } = workbook.usedSize(sheetId)
  return Array.from({ length: height }, (_, row) => Array.from({ length: width }, (_, col) =>
    formatCellValue(workbook.getCellValue(sheetId, [col, row]), workbook.getStyle(sheetId, [col, row]))))
}

function extensionOf(path: string): string {
  return path.split('.').pop()?.toLowerCase() ?? ''
}

async function encode(workbook: Workbook, format: string): Promise<Uint8Array> {
  if (!isDelimitedFormat(format)) return writeXlsx(workbook.toData())
  const first = workbook.sheets()[0]
  return writeDelimited(first === undefined ? [] : displayRows(workbook, first.id), format)
}

/**
 * Saves a workbook tab. In-place saves go straight to disk; a dialog opens for Save As, for untitled
 * documents, for formats this app cannot write (.xlsm), for multi-sheet workbooks kept as delimited
 * text (.csv, .tsv), and when the import report says a save would drop content, so the original is
 * never silently overwritten.
 */
export async function saveWorkbook(documentId: string, saveAs: boolean): Promise<boolean> {
  const document = useDocumentsStore.getState().documents.find(candidate => candidate.id === documentId)
  const workbook = workbookFor(documentId)
  if (document === undefined || workbook === undefined) return false

  const report = useFidelityStore.getState().reports[documentId]
  const losesContent = report?.severity === 'dropped'
  const multiSheet = workbook.sheets().length > 1
  const multiSheetText = isDelimitedFormat(document.extension) && multiSheet
  const needsDialog = saveAs || document.path === undefined || !isWritableExtension(document.extension) || losesContent || multiSheetText
  let path = document.path
  if (needsDialog) {
    const format = isDelimitedFormat(document.extension) && !multiSheetText ? document.extension : 'xlsx'
    const suffix = losesContent && !saveAs ? ' (edited)' : ''
    path = await fileService.chooseSavePath({
      defaultName: `${document.name}${suffix}.${format}`,
      // Delimited text holds one sheet; a multi-sheet workbook is only offered .xlsx (Export Sheet as CSV covers one sheet).
      extensions: multiSheet ? ['xlsx'] : [format, ...findApplication('sheets').saves.filter(extension => extension !== format)],
    })
    if (path === undefined) return false
  }
  if (path === undefined) return false

  const format = extensionOf(path)
  const point = workbook.savePoint()
  await fileService.write(path, await encode(workbook, format))
  const file = await fileService.describe(path)
  workbook.markSaved(point)

  const cached = cache.get(documentId)
  const newId = useDocumentsStore.getState().setSaved(documentId, file)
  if (cached !== undefined && newId !== documentId) {
    cache.delete(documentId)
    cached.documentId = newId
    cache.set(newId, cached)
    const previousReport = useFidelityStore.getState().reports[documentId]
    useFidelityStore.getState().forget(documentId)
    // A file this app just wrote has nothing unrepresented in it.
    if (previousReport !== undefined) useFidelityStore.getState().publish(newId, createImportReport(file.name, []))
    emit()
  }
  await recordRecent(file)
  const first = workbook.sheets()[0]
  if (first !== undefined) await recordPreview(file.path, renderPreview(workbook, first.id))
  toast.success(`Saved ${file.name}`)
  if (isDelimitedFormat(format)) {
    toast.info(`${format.toUpperCase()} keeps values only`, { description: 'Formulas, formatting and other sheets are saved in .xlsx.' })
  }
  return true
}

/** Writes the active sheet's visible values to a CSV chosen by the user, without changing the document. */
export async function exportSheetCsv(documentId: string, sheetId: number, sheetName: string): Promise<void> {
  const document = useDocumentsStore.getState().documents.find(candidate => candidate.id === documentId)
  const workbook = workbookFor(documentId)
  if (document === undefined || workbook === undefined) return
  const path = await fileService.chooseSavePath({ defaultName: `${document.name} - ${sheetName}.csv`, extensions: ['csv'] })
  if (path === undefined) return
  await fileService.write(path, writeDelimited(displayRows(workbook, sheetId), 'csv'))
  toast.success(`Exported ${path.split('/').pop() ?? 'CSV'}`)
}

/** The sheet's printable table; says so and resolves undefined when there is nothing to print. */
function printableTable(workbook: Workbook, sheetId: number): SheetPrintTable | undefined {
  const table = sheetPrintTable(workbook, sheetId)
  if (table === undefined) toast.info('Nothing to print', { description: 'The visible cells of this sheet are empty.' })
  return table
}

/** Opens the print dialog for one sheet. */
export async function printSheet(workbook: Workbook, sheetId: number): Promise<void> {
  const table = printableTable(workbook, sheetId)
  if (table !== undefined) await openPrintDialog(paper => sheetPrintout(table, paper))
}

/** Writes one sheet, as it prints, to a PDF chosen by the user, and offers to open it in Hanko. */
export async function exportSheetPdf(documentId: string, sheetId: number): Promise<void> {
  const document = useDocumentsStore.getState().documents.find(candidate => candidate.id === documentId)
  const workbook = workbookFor(documentId)
  const table = workbook === undefined ? undefined : printableTable(workbook, sheetId)
  if (document === undefined || table === undefined) return
  const path = await fileService.chooseSavePath({ defaultName: `${document.name}.pdf`, extensions: ['pdf'] })
  if (path === undefined) return
  const bytes = await printToPdf(paper => sheetPrintout(table, paper))
  if (bytes === undefined) return
  await fileService.write(path, bytes)
  const file = await fileService.describe(path)
  toast.success('Exported as PDF', {
    description: file.name,
    action: { label: 'Open', onClick: () => void openFiles([file]) },
  })
}
