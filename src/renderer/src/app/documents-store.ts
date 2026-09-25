import { create } from 'zustand'
import { applicationForExtension, findApplication, type EditorApplicationId } from '@shared/applications'
import type { FileReference } from '@shared/shell'

/** The pinned first tab. Document ids are paths or `untitled:<n>`, so they never collide with it. */
export const homeTabId = 'home'

/** One document tab. Path-backed documents use their path as id; untitled ones use `untitled:<n>`. */
export interface OpenDocument {
  readonly id: string
  readonly name: string
  readonly extension: string
  readonly kind: EditorApplicationId
  readonly path?: string
  readonly dirty: boolean
}

interface DocumentsState {
  readonly documents: readonly OpenDocument[]
  /** `homeTabId` or a document id. */
  readonly activeId: string
  openFile(file: FileReference): string
  /** Adds a tab from last session without taking focus. */
  restoreFile(file: FileReference): void
  openUntitled(kind: EditorApplicationId): string
  select(id: string): void
  /** Selects by position, Home being 0; a position past the end selects the last tab. */
  selectIndex(index: number): void
  /** Moves the selection left or right, wrapping around through Home. */
  selectRelative(step: -1 | 1): void
  reorder(activeId: string, overId: string): void
  remove(id: string): void
  setDirty(id: string, dirty: boolean): void
  /** Re-keys a document after Save As or a first save, keeping its tab position. */
  setSaved(id: string, file: FileReference): string
}

function nameWithoutExtension(name: string): string {
  return name.replace(/\.[^.]+$/, '')
}

function untitledName(documents: readonly OpenDocument[]): string {
  const taken = new Set(documents.map(document => document.name))
  if (!taken.has('Untitled')) return 'Untitled'
  let index = 2
  while (taken.has(`Untitled ${index}`)) index += 1
  return `Untitled ${index}`
}

function fileDocument(file: FileReference): OpenDocument {
  const kind = applicationForExtension(file.extension)
  if (kind === undefined) throw new Error(`.${file.extension || 'unknown'} files cannot be opened.`)
  return { id: file.path, name: nameWithoutExtension(file.name), extension: file.extension, kind, path: file.path, dirty: false }
}

/** A new tab joins the end of its own kind's group, so spreadsheets stay together and PDFs stay together. */
function insertGrouped(documents: readonly OpenDocument[], document: OpenDocument): readonly OpenDocument[] {
  const last = documents.findLastIndex(candidate => candidate.kind === document.kind)
  const at = last < 0 ? documents.length : last + 1
  return [...documents.slice(0, at), document, ...documents.slice(at)]
}

let untitledCounter = 0

export const useDocumentsStore = create<DocumentsState>((set, get) => ({
  documents: [],
  activeId: homeTabId,
  openFile: file => {
    const existing = get().documents.find(document => document.path === file.path)
    if (existing !== undefined) {
      set({ activeId: existing.id })
      return existing.id
    }
    const document = fileDocument(file)
    set(state => ({ documents: insertGrouped(state.documents, document), activeId: document.id }))
    return document.id
  },
  restoreFile: file => {
    if (get().documents.some(document => document.path === file.path)) return
    set(state => ({ documents: [...state.documents, fileDocument(file)] }))
  },
  openUntitled: kind => {
    untitledCounter += 1
    const document: OpenDocument = {
      id: `untitled:${untitledCounter}`,
      name: untitledName(get().documents),
      extension: findApplication(kind).opens[0] ?? '',
      kind,
      dirty: false,
    }
    set(state => ({ documents: insertGrouped(state.documents, document), activeId: document.id }))
    return document.id
  },
  select: id => set({ activeId: id }),
  selectIndex: index => {
    const ids = [homeTabId, ...get().documents.map(document => document.id)]
    const id = ids[Math.min(index, ids.length - 1)]
    if (id !== undefined) set({ activeId: id })
  },
  selectRelative: step => {
    const ids = [homeTabId, ...get().documents.map(document => document.id)]
    const current = ids.indexOf(get().activeId)
    const id = ids[(current + step + ids.length) % ids.length]
    if (id !== undefined) set({ activeId: id })
  },
  reorder: (activeId, overId) => set(state => {
    const from = state.documents.findIndex(document => document.id === activeId)
    const to = state.documents.findIndex(document => document.id === overId)
    if (from < 0 || to < 0 || from === to) return state
    const documents = [...state.documents]
    const [moved] = documents.splice(from, 1)
    if (moved !== undefined) documents.splice(to, 0, moved)
    return { documents }
  }),
  remove: id => set(state => {
    const index = state.documents.findIndex(document => document.id === id)
    const documents = state.documents.filter(document => document.id !== id)
    if (state.activeId !== id) return { documents }
    // Closing the active tab selects its right neighbour, or the left one at the end, like Safari.
    const next = documents[Math.min(index, documents.length - 1)]
    return { documents, activeId: next?.id ?? homeTabId }
  }),
  setDirty: (id, dirty) => set(state => ({
    documents: state.documents.map(document =>
      document.id === id && document.dirty !== dirty ? { ...document, dirty } : document),
  })),
  setSaved: (id, file) => {
    set(state => ({
      documents: state.documents.map(document => document.id === id ? fileDocument(file) : document),
      activeId: state.activeId === id ? file.path : state.activeId,
    }))
    return file.path
  },
}))

export function activeDocument(state: Pick<DocumentsState, 'documents' | 'activeId'>): OpenDocument | undefined {
  return state.documents.find(document => document.id === state.activeId)
}
