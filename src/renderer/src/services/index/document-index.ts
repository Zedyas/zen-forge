import { openDB, type DBSchema } from 'idb'
import type { FileReference } from '@shared/shell'

export interface RecentDocument extends FileReference {
  readonly openedAt: number
  /** A small PNG data URL of the first page or the top-left cells, shown on Home. */
  readonly preview?: string
}

interface DocumentIndexSchema extends DBSchema {
  recents: {
    key: string
    value: RecentDocument
    indexes: { 'by-opened-at': number }
  }
}

const databasePromise = openDB<DocumentIndexSchema>('document-index', 1, {
  upgrade(database) {
    const recents = database.createObjectStore('recents', { keyPath: 'path' })
    recents.createIndex('by-opened-at', 'openedAt')
  },
})

const recentEvents = new BroadcastChannel('document-index-recents')

/** Inserts or refreshes a disposable index entry for a document opened from disk, keeping its preview. */
export async function recordRecent(file: FileReference): Promise<void> {
  const database = await databasePromise
  const existing = await database.get('recents', file.path)
  await database.put('recents', { ...file, openedAt: Date.now(), preview: existing?.preview })
  recentEvents.postMessage('changed')
}

export async function recordPreview(path: string, preview: string): Promise<void> {
  const database = await databasePromise
  const existing = await database.get('recents', path)
  if (existing === undefined || existing.preview === preview) return
  await database.put('recents', { ...existing, preview })
  recentEvents.postMessage('changed')
}

export async function forgetRecent(path: string): Promise<void> {
  const database = await databasePromise
  await database.delete('recents', path)
  recentEvents.postMessage('changed')
}

export async function listRecents(query = ''): Promise<readonly RecentDocument[]> {
  const database = await databasePromise
  const documents = await database.getAllFromIndex('recents', 'by-opened-at')
  const normalizedQuery = query.trim().toLocaleLowerCase()
  return documents
    .reverse()
    .filter(document => normalizedQuery.length === 0 || document.name.toLocaleLowerCase().includes(normalizedQuery))
    .slice(0, 40)
}

/** Refreshes Home when another window updates the shared index. */
export function watchRecents(listener: () => void): () => void {
  const handleMessage = (): void => listener()
  recentEvents.addEventListener('message', handleMessage)
  return () => recentEvents.removeEventListener('message', handleMessage)
}
