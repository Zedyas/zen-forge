import { create } from 'zustand'
import { useDocumentsStore } from '../../app/documents-store'
import { findElement, findSlide, type Presentation, type Slide, type SlideElement } from './model'

export interface ReadySlides {
  readonly status: 'ready'
  readonly past: readonly Presentation[]
  readonly present: Presentation
  readonly future: readonly Presentation[]
  readonly saved: Presentation
  /** The slide shown on the canvas. */
  readonly slideId: string
  /** The selected elements on that slide, in the order they were selected. */
  readonly selection: readonly string[]
  /** The text box, shape or table whose text is being typed into. */
  readonly editingId?: string
  /** The current cell of the selected table: the one being typed into, or that row and column actions act on. */
  readonly cell?: TableCellAddress
  /** Index into `present.slides` while the slideshow plays. */
  readonly playing?: number
}

export interface TableCellAddress {
  readonly row: number
  readonly column: number
}

export type SlidesDocument =
  | { readonly status: 'loading' }
  | { readonly status: 'error'; readonly message: string }
  | ReadySlides

interface SlidesStore {
  readonly documents: Readonly<Record<string, SlidesDocument>>
}

export const useSlidesStore = create<SlidesStore>(() => ({ documents: {} }))

const historyLimit = 100

function syncDirty(id: string, document: ReadySlides): void {
  useDocumentsStore.getState().setDirty(id, document.present !== document.saved)
}

/**
 * Text editors (a text box, a table cell, the notes) keep what is typed until they close. The tab is
 * marked unsaved at the first keystroke, so closing the window or quitting asks before it is lost.
 */
export function markTyping(id: string): void {
  useDocumentsStore.getState().setDirty(id, true)
}

/** After an editor wrote its text (or had nothing to write), the unsaved mark follows the presentation again. */
export function typingDone(id: string): void {
  const document = readySlides(id)
  if (document !== undefined) syncDirty(id, document)
}

export function setSlidesDocument(id: string, document: SlidesDocument): void {
  useSlidesStore.setState(state => ({ documents: { ...state.documents, [id]: document } }))
}

export function readySlides(id: string): ReadySlides | undefined {
  const document = useSlidesStore.getState().documents[id]
  return document?.status === 'ready' ? document : undefined
}

export function readyDocument(presentation: Presentation): ReadySlides {
  return { status: 'ready', past: [], present: presentation, future: [], saved: presentation, slideId: presentation.slides[0]?.id ?? '', selection: [] }
}

export function currentSlide(document: ReadySlides): Slide | undefined {
  return findSlide(document.present, document.slideId) ?? document.present.slides[0]
}

/** The selected elements, back to front. */
export function selectedElements(document: ReadySlides): SlideElement[] {
  const ids = new Set(document.selection)
  return currentSlide(document)?.elements.filter(element => ids.has(element.id)) ?? []
}

/** The selected element when exactly one is selected. */
export function selectedElement(document: ReadySlides): SlideElement | undefined {
  return document.selection.length === 1 ? findElement(currentSlide(document), document.selection[0]) : undefined
}

/** Changes view state (current slide, selection, editing, playing) without touching undo history. */
export function updateSlides(id: string, change: (document: ReadySlides) => Partial<ReadySlides>): void {
  const document = readySlides(id)
  if (document !== undefined) setSlidesDocument(id, { ...document, ...change(document) })
}

/** Keeps the current slide, selection and cell pointing at things that still exist after an edit, undo or redo. */
function settle(document: ReadySlides): ReadySlides {
  const slide = currentSlide(document)
  const selection = document.selection.filter(id => findElement(slide, id) !== undefined)
  const only = selection.length === 1 ? findElement(slide, selection[0]) : undefined
  const cell = only?.kind === 'table' && document.cell !== undefined && document.cell.row < only.rows.length && document.cell.column < only.columns.length
    ? document.cell
    : undefined
  const editingId = document.editingId !== undefined && selection.includes(document.editingId) ? document.editingId : undefined
  return { ...document, slideId: slide?.id ?? '', selection, cell, editingId }
}

/** Records one undoable edit. */
export function commitSlides(id: string, change: (presentation: Presentation) => Presentation): void {
  const document = readySlides(id)
  if (document === undefined) return
  const present = change(document.present)
  if (present === document.present) return
  const next = settle({ ...document, past: [...document.past, document.present].slice(-historyLimit), present, future: [] })
  setSlidesDocument(id, next)
  syncDirty(id, next)
}

export function undoSlides(id: string): void {
  const document = readySlides(id)
  const previous = document?.past.at(-1)
  if (document === undefined || previous === undefined) return
  const next = settle({ ...document, past: document.past.slice(0, -1), present: previous, future: [...document.future, document.present], editingId: undefined })
  setSlidesDocument(id, next)
  syncDirty(id, next)
}

export function redoSlides(id: string): void {
  const document = readySlides(id)
  const following = document?.future.at(-1)
  if (document === undefined || following === undefined) return
  const next = settle({ ...document, past: [...document.past, document.present], present: following, future: document.future.slice(0, -1), editingId: undefined })
  setSlidesDocument(id, next)
  syncDirty(id, next)
}

export function markSlidesSaved(id: string, presentation: Presentation): void {
  const document = readySlides(id)
  if (document === undefined) return
  const next = { ...document, saved: presentation }
  setSlidesDocument(id, next)
  syncDirty(id, next)
}

/** Moves a document's state to a new id after Save As changed its path. */
export function rekeySlides(from: string, to: string): void {
  if (from === to) return
  useSlidesStore.setState(state => {
    const documents = { ...state.documents }
    const document = documents[from]
    delete documents[from]
    if (document !== undefined) documents[to] = document
    return { documents }
  })
}

export function releaseSlides(id: string): void {
  useSlidesStore.setState(state => {
    const documents = { ...state.documents }
    delete documents[id]
    return { documents }
  })
}
