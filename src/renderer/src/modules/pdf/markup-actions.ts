import { newId, type Markup, type PageItem, type Snapshot } from './model'
import { commitPdf } from './pdf-store'
import { select } from './tool-store'

function onPage(snapshot: Snapshot, pageKey: string, change: (page: PageItem) => PageItem): Snapshot {
  return { ...snapshot, pages: snapshot.pages.map(page => page.key === pageKey ? change(page) : page) }
}

/** Adds a markup as one undoable step and selects it, so its properties show in the inspector. */
export function addMarkup(documentId: string, pageKey: string, markup: Markup): string {
  const id = newId()
  commitPdf(documentId, snapshot => onPage(snapshot, pageKey, page => ({ ...page, markups: [...page.markups, { id, markup }] })))
  select({ pageKey, markupId: id })
  return id
}

export function updateMarkup(documentId: string, pageKey: string, markupId: string, markup: Markup): void {
  commitPdf(documentId, snapshot => onPage(snapshot, pageKey, page => ({
    ...page,
    markups: page.markups.map(placed => placed.id === markupId ? { id: markupId, markup } : placed),
  })))
}

export function removeMarkup(documentId: string, pageKey: string, markupId: string): void {
  commitPdf(documentId, snapshot => onPage(snapshot, pageKey, page => ({
    ...page,
    markups: page.markups.filter(placed => placed.id !== markupId),
  })))
  select(undefined)
}

export function findMarkup(snapshot: Snapshot, pageKey: string, markupId: string): Markup | undefined {
  return snapshot.pages.find(page => page.key === pageKey)?.markups.find(placed => placed.id === markupId)?.markup
}
