import { toast } from 'sonner'
import type { CommandId } from '@shared/commands'
import { createImportReport } from '@shared/fidelity'
import { useDocumentsStore } from '../../app/documents-store'
import { fileService } from '../../services/file/IpcFileService'
import { openFiles } from '../../app/document-actions'
import { useFidelityStore } from '../../services/fidelity/fidelity-store'
import { recordPreview, recordRecent } from '../../services/index/document-index'
import { askConfirm } from './ConfirmDialog'
import { inspectPdf, listFormFields, redactPages, savePdf, type RedactOptions } from './engine'
import { openFind } from './find-store'
import { isRedaction, newId, toPageRef, turnPage, type PageItem, type Size, type Snapshot } from './model'
import { closePdfJs, openPdfJs, renderPage } from './pdfjs'
import { printPdfDocument } from './print'
import {
  commitPdf,
  markPdfSaved,
  pageSize,
  readyPdf,
  redoPdf,
  rekeyPdf,
  releasePdf,
  setPdfDocument,
  sizeKey,
  undoPdf,
  updatePdf,
  type ReadyPdf,
} from './pdf-store'
import { select } from './tool-store'

export const zoomLevels: readonly number[] = [0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4]

function errorMessage(error: unknown, fallback: string): string {
  return error instanceof Error && error.message !== '' ? error.message : fallback
}

/** Reads each page's displayed size (own /Rotate applied) for layout before anything renders. */
async function measureSource(bytes: Uint8Array, source: number): Promise<Map<string, Size>> {
  const document = await openPdfJs(bytes)
  const sizes = new Map<string, Size>()
  for (let index = 0; index < document.numPages; index += 1) {
    const viewport = (await document.getPage(index + 1)).getViewport({ scale: 1 })
    sizes.set(sizeKey(source, index), { width: viewport.width, height: viewport.height })
  }
  return sizes
}

function pagesOf(source: number, count: number): PageItem[] {
  return Array.from({ length: count }, (_, index) => ({ key: newId(), source, index, rotation: 0, markups: [] }))
}

function friendlyLoadError(error: unknown): string {
  if (error instanceof Error && error.name === 'PasswordException') {
    return 'This PDF is password-protected. Opening protected PDFs is not supported yet.'
  }
  if (error instanceof Error && error.name === 'InvalidPDFException') {
    return 'This file is not a valid PDF, or it is damaged.'
  }
  return errorMessage(error, 'The file could not be read.')
}

/** The latest load of each document. A load that is no longer the latest, because its tab closed or reopened, drops its result. */
const loads = new Map<string, symbol>()

/** Loads a path-backed document into the store and publishes its import report. */
export async function loadPdfDocument(id: string, path: string, fileName: string): Promise<void> {
  const load = Symbol(id)
  loads.set(id, load)
  setPdfDocument(id, { status: 'loading' })
  let bytes: Uint8Array | undefined
  try {
    bytes = await fileService.read(path)
    const inspection = await inspectPdf(bytes)
    const sizes = await measureSource(bytes, 0)
    const fields = inspection.hasForm ? await listFormFields(bytes) : []
    const snapshot: Snapshot = { pages: pagesOf(0, sizes.size), formValues: {}, flattenForm: false }
    if (loads.get(id) !== load) {
      void closePdfJs(bytes)
      return
    }
    setPdfDocument(id, {
      status: 'ready',
      sources: [bytes],
      sizes,
      fields,
      past: [],
      present: snapshot,
      future: [],
      saved: snapshot,
      zoom: 1,
      fitted: false,
      currentPage: 0,
    })
    useFidelityStore.getState().publish(id, createImportReport(fileName, inspection.findings))
    void recordPagePreview(path, bytes)
  } catch (error) {
    if (bytes !== undefined) void closePdfJs(bytes)
    if (loads.get(id) === load) setPdfDocument(id, { status: 'error', message: friendlyLoadError(error) })
  }
}

/** Renders page 1 small for the Home recents grid, from the pdf.js copy `bytes` already has or a new one. */
async function recordPagePreview(path: string, bytes: Uint8Array): Promise<void> {
  try {
    const document = await openPdfJs(bytes)
    const page = await document.getPage(1)
    const width = page.getViewport({ scale: 1 }).width
    const canvas = window.document.createElement('canvas')
    const task = await renderPage(document, 0, canvas, { scale: 220 / width, extraRotation: 0, forms: 'print', pixelRatio: 1 })
    await task.promise
    await recordPreview(path, canvas.toDataURL('image/png'))
  } catch {
    // A missing preview only costs the thumbnail on Home.
  }
}

export function releasePdfDocument(id: string): void {
  loads.delete(id)
  const document = readyPdf(id)
  document?.sources.forEach(source => void closePdfJs(source))
  releasePdf(id)
  useFidelityStore.getState().forget(id)
}

/** Builds the output bytes for some pages: the normal save, then redaction of the pages that have boxes. */
async function buildOutput(document: ReadyPdf, pages: readonly PageItem[], redaction: RedactOptions): Promise<Uint8Array> {
  const snapshot = document.present
  const saved = await savePdf({
    sources: document.sources,
    pages: pages.map(toPageRef),
    formValues: snapshot.formValues,
    flattenForm: snapshot.flattenForm,
  })
  const requests = pages.flatMap((item, index) => {
    const boxes = item.markups.flatMap(placed => isRedaction(placed.markup) ? [placed.markup] : [])
    return boxes.length === 0 ? [] : [{ index, boxes }]
  })
  if (requests.length === 0) return saved
  return redactPages(saved, requests, redaction)
}

function hasRedactions(pages: readonly PageItem[]): boolean {
  return pages.some(item => item.markups.some(placed => isRedaction(placed.markup)))
}

/** Asks before redacting. Resolves the redaction options, or undefined when cancelled. */
async function confirmRedaction(pages: readonly PageItem[]): Promise<RedactOptions | undefined> {
  if (!hasRedactions(pages)) return { removeHiddenInformation: false }
  const answer = await askConfirm({
    title: 'Apply redactions?',
    message: 'Everything under the redaction boxes is permanently removed from the saved file: text, the covered parts of images, and drawings. The rest of each page stays as it is.',
    option: {
      label: 'Also remove hidden information',
      detail: 'Document title and author, bookmarks, attached files, comments, hidden layers, and scripts.',
      checked: true,
    },
    actions: [{ label: 'Redact and save', value: 'redact', primary: true }],
  })
  return answer.value === 'redact' ? { removeHiddenInformation: answer.optionChecked } : undefined
}

/** Saves in place, or asks for a path. Keeps the edit model and undo history; the originals stay in memory. */
export async function savePdfDocument(id: string, saveAs: boolean): Promise<boolean> {
  // A text box or form field being typed in commits on blur; do that first so the save includes it.
  if (window.document.activeElement instanceof HTMLElement) window.document.activeElement.blur()
  const document = readyPdf(id)
  const open = useDocumentsStore.getState().documents.find(candidate => candidate.id === id)
  if (document === undefined || open === undefined) return false

  const redaction = await confirmRedaction(document.present.pages)
  if (redaction === undefined) return false

  let chooseNewPath = saveAs || open.path === undefined
  if (!chooseNewPath && useFidelityStore.getState().reports[id]?.severity === 'dropped') {
    const { value } = await askConfirm({
      title: 'Save over the original?',
      message: 'This PDF contains content that saving removes (listed in the import report). Saving a copy keeps the original file intact.',
      actions: [{ label: 'Overwrite', value: 'overwrite' }, { label: 'Save a copy', value: 'copy', primary: true }],
    })
    if (value === 'cancel') return false
    chooseNewPath = value === 'copy'
  }

  const path = chooseNewPath
    ? await fileService.chooseSavePath({ defaultName: `${open.name}.pdf`, extensions: ['pdf'] })
    : open.path
  if (path === undefined) return false

  const snapshot = document.present
  const bytes = await buildOutput(document, snapshot.pages, redaction)
  await fileService.write(path, bytes)
  const file = await fileService.describe(path)
  const nextId = useDocumentsStore.getState().setSaved(id, file)
  rekeyPdf(id, nextId)
  markPdfSaved(nextId, snapshot)
  if (nextId !== id) {
    const report = useFidelityStore.getState().reports[id]
    useFidelityStore.getState().forget(id)
    if (report !== undefined) useFidelityStore.getState().publish(nextId, createImportReport(file.name, (await inspectPdf(bytes)).findings))
  }
  await recordRecent(file)
  // The saved bytes are not one of this document's sources, so their pdf.js copy goes after the preview.
  void recordPagePreview(path, bytes).finally(() => void closePdfJs(bytes))
  toast.success('Saved', { description: file.name })
  return true
}

function activeId(): string | undefined {
  return useDocumentsStore.getState().activeId
}

function currentItem(document: ReadyPdf): PageItem | undefined {
  return document.present.pages[document.currentPage]
}

export function rotateCurrentPage(id: string, turn: 1 | -1): void {
  const document = readyPdf(id)
  const item = document === undefined ? undefined : currentItem(document)
  if (document === undefined || item === undefined) return
  const size = pageSize(document, item)
  commitPdf(id, snapshot => ({
    ...snapshot,
    pages: snapshot.pages.map(page => page.key === item.key ? turnPage(page, size, turn) : page),
  }))
}

export async function deleteCurrentPage(id: string): Promise<void> {
  const document = readyPdf(id)
  const item = document === undefined ? undefined : currentItem(document)
  if (document === undefined || item === undefined) return
  if (document.present.pages.length === 1) {
    toast.error('A PDF needs at least one page.', { description: 'Delete the file in Finder to remove it entirely.' })
    return
  }
  select(undefined)
  commitPdf(id, snapshot => ({ ...snapshot, pages: snapshot.pages.filter(page => page.key !== item.key) }))
  updatePdf(id, current => ({ currentPage: Math.min(current.currentPage, current.present.pages.length - 1) }))
}

export function movePage(id: string, fromKey: string, toKey: string): void {
  commitPdf(id, snapshot => {
    const from = snapshot.pages.findIndex(page => page.key === fromKey)
    const to = snapshot.pages.findIndex(page => page.key === toKey)
    if (from < 0 || to < 0 || from === to) return snapshot
    const pages = [...snapshot.pages]
    const [moved] = pages.splice(from, 1)
    if (moved !== undefined) pages.splice(to, 0, moved)
    return { ...snapshot, pages }
  })
}

/** Adds another PDF as a source and inserts all its pages after the current page. */
export async function insertPagesFromFile(id: string): Promise<void> {
  const document = readyPdf(id)
  if (document === undefined) return
  const [file] = await fileService.chooseFiles({ extensions: ['pdf'] })
  if (file === undefined) return
  const bytes = await fileService.read(file.path)
  const source = document.sources.length
  const sizes = await measureSource(bytes, source)
  updatePdf(id, current => ({ sources: [...current.sources, bytes], sizes: new Map([...current.sizes, ...sizes]) }))
  const inserted = pagesOf(source, sizes.size)
  const at = document.currentPage + 1
  commitPdf(id, snapshot => ({ ...snapshot, pages: [...snapshot.pages.slice(0, at), ...inserted, ...snapshot.pages.slice(at)] }))
  toast.success(`Inserted ${inserted.length} ${inserted.length === 1 ? 'page' : 'pages'}`, { description: file.name })
}

/** Writes the current page, with its edits and redactions, to a new PDF and offers to open it. */
export async function extractCurrentPage(id: string): Promise<void> {
  const document = readyPdf(id)
  const item = document === undefined ? undefined : currentItem(document)
  const open = useDocumentsStore.getState().documents.find(candidate => candidate.id === id)
  if (document === undefined || item === undefined || open === undefined) return
  const redaction = await confirmRedaction([item])
  if (redaction === undefined) return
  const path = await fileService.chooseSavePath({
    defaultName: `${open.name} (page ${document.currentPage + 1}).pdf`,
    extensions: ['pdf'],
  })
  if (path === undefined) return
  await fileService.write(path, await buildOutput(document, [item], redaction))
  const file = await fileService.describe(path)
  toast.success('Page extracted', {
    description: file.name,
    action: { label: 'Open', onClick: () => void openFiles([file]) },
  })
}

export function setZoom(id: string, zoom: number): void {
  updatePdf(id, () => ({ zoom: Math.min(4, Math.max(0.25, Math.round(zoom * 100) / 100)) }))
}

function stepZoom(id: string, direction: 1 | -1): void {
  const document = readyPdf(id)
  if (document === undefined) return
  const next = direction === 1
    ? zoomLevels.find(level => level > document.zoom + 0.001)
    : [...zoomLevels].reverse().find(level => level < document.zoom - 0.001)
  if (next !== undefined) setZoom(id, next)
}

export function isTextEntry(target: EventTarget | null): boolean {
  return target instanceof HTMLElement && target.matches('input, textarea, select, [contenteditable="true"]')
}

/** Editor commands from the menu, keyboard and toolbar, applied to the active document. */
export async function runPdfCommand(command: CommandId): Promise<void> {
  const id = activeId()
  if (id === undefined) return
  try {
    switch (command) {
      case 'edit.undo':
        // A focused field keeps its own text undo; the menu accelerator would otherwise swallow it.
        if (isTextEntry(window.document.activeElement)) window.document.execCommand('undo')
        else undoPdf(id)
        return
      case 'edit.redo':
        if (isTextEntry(window.document.activeElement)) window.document.execCommand('redo')
        else redoPdf(id)
        return
      case 'edit.find':
        return openFind()
      case 'view.zoomIn':
        return stepZoom(id, 1)
      case 'view.zoomOut':
        return stepZoom(id, -1)
      case 'view.actualSize':
        return setZoom(id, 1)
      case 'page.rotateLeft':
        return rotateCurrentPage(id, -1)
      case 'page.rotateRight':
        return rotateCurrentPage(id, 1)
      case 'page.delete':
        return await deleteCurrentPage(id)
      case 'page.insert':
        return await insertPagesFromFile(id)
      case 'page.extract':
        return await extractCurrentPage(id)
      case 'file.print':
        return await printPdfDocument(id)
      default:
        return
    }
  } catch (error) {
    toast.error('Could not complete that', { description: errorMessage(error, 'The PDF could not be changed.') })
  }
}
