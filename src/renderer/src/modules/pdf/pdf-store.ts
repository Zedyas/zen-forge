import { create } from 'zustand'
import { useDocumentsStore } from '../../app/documents-store'
import type { FormField } from './engine'
import { changesSinceSave, rotatedSize, type PageItem, type Size, type Snapshot } from './model'

/** Displayed size of a source page at 100% with its own /Rotate applied, before any user rotation. */
export type PageSizes = ReadonlyMap<string, Size>

export function sizeKey(source: number, index: number): string {
  return `${source}:${index}`
}

export interface ReadyPdf {
  readonly status: 'ready'
  readonly sources: readonly Uint8Array[]
  readonly sizes: PageSizes
  readonly fields: readonly FormField[]
  readonly past: readonly Snapshot[]
  readonly present: Snapshot
  readonly future: readonly Snapshot[]
  readonly saved: Snapshot
  readonly zoom: number
  /** Set once the zoom has been fitted to the desk width; later visits keep the user's zoom. */
  readonly fitted: boolean
  /** Index into `present.pages` of the page nearest the top of the view. */
  readonly currentPage: number
}

export type PdfDocument =
  | { readonly status: 'loading' }
  | { readonly status: 'error'; readonly message: string }
  | ReadyPdf

interface PdfStore {
  readonly documents: Readonly<Record<string, PdfDocument>>
}

export const usePdfStore = create<PdfStore>(() => ({ documents: {} }))

const historyLimit = 100

function syncDirty(id: string, document: ReadyPdf): void {
  useDocumentsStore.getState().setDirty(id, document.present !== document.saved)
}

/** A page's displayed size in points at 100%, with its own /Rotate and the user's rotation applied. */
export function pageSize(document: ReadyPdf, item: PageItem): Size {
  return rotatedSize(document.sizes.get(sizeKey(item.source, item.index)) ?? { width: 612, height: 792 }, item.rotation)
}

export function setPdfDocument(id: string, document: PdfDocument): void {
  usePdfStore.setState(state => ({ documents: { ...state.documents, [id]: document } }))
}

export function readyPdf(id: string): ReadyPdf | undefined {
  const document = usePdfStore.getState().documents[id]
  return document?.status === 'ready' ? document : undefined
}

/** Changes view-only fields (zoom, current page, sources, sizes) without touching undo history. */
export function updatePdf(id: string, change: (document: ReadyPdf) => Partial<ReadyPdf>): void {
  const document = readyPdf(id)
  if (document !== undefined) setPdfDocument(id, { ...document, ...change(document) })
}

/** Records one undoable edit. */
export function commitPdf(id: string, change: (snapshot: Snapshot) => Snapshot): void {
  const document = readyPdf(id)
  if (document === undefined) return
  const present = change(document.present)
  if (present === document.present) return
  const next: ReadyPdf = { ...document, past: [...document.past, document.present].slice(-historyLimit), present, future: [] }
  setPdfDocument(id, next)
  syncDirty(id, next)
}

export function undoPdf(id: string): void {
  const document = readyPdf(id)
  const previous = document?.past.at(-1)
  if (document === undefined || previous === undefined) return
  const next: ReadyPdf = { ...document, past: document.past.slice(0, -1), present: previous, future: [...document.future, document.present] }
  setPdfDocument(id, next)
  syncDirty(id, next)
}

export function redoPdf(id: string): void {
  const document = readyPdf(id)
  const following = document?.future.at(-1)
  if (document === undefined || following === undefined) return
  const next: ReadyPdf = { ...document, past: [...document.past, document.present], present: following, future: document.future.slice(0, -1) }
  setPdfDocument(id, next)
  syncDirty(id, next)
}

export function markPdfSaved(id: string, snapshot: Snapshot): void {
  const document = readyPdf(id)
  if (document === undefined) return
  const next = { ...document, saved: snapshot }
  setPdfDocument(id, next)
  syncDirty(id, next)
}

/** Moves a document's state to a new id after Save As changed its path. */
export function rekeyPdf(from: string, to: string): void {
  if (from === to) return
  usePdfStore.setState(state => {
    const documents = { ...state.documents }
    const document = documents[from]
    delete documents[from]
    if (document !== undefined) documents[to] = document
    return { documents }
  })
}

export function releasePdf(id: string): void {
  usePdfStore.setState(state => {
    const documents = { ...state.documents }
    delete documents[id]
    return { documents }
  })
}

export function unsavedChanges(document: ReadyPdf): number {
  return changesSinceSave(document.past, document.present, document.future, document.saved)
}
