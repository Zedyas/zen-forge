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
const localListeners = new Set<() => void>()

/** A channel does not deliver to its own sender, so this window's listeners are called directly. */
function announceChange(): void {
  recentEvents.postMessage('changed')
  localListeners.forEach(listener => listener())
}

/** Inserts or refreshes a disposable index entry for a document opened from disk, keeping its preview. */
export async function recordRecent(file: FileReference): Promise<void> {
  const database = await databasePromise
  // Reading and writing in one transaction, so a preview saved meanwhile does not write the old time back.
  const transaction = database.transaction('recents', 'readwrite')
  const existing = await transaction.store.get(file.path)
  await transaction.store.put({ ...file, openedAt: Date.now(), preview: existing?.preview })
  await transaction.done
  announceChange()
}

export async function recordPreview(path: string, preview: string): Promise<void> {
  const database = await databasePromise
  const transaction = database.transaction('recents', 'readwrite')
  const existing = await transaction.store.get(path)
  if (existing === undefined || existing.preview === preview) return
  await transaction.store.put({ ...existing, preview })
  await transaction.done
  announceChange()
}

export async function forgetRecent(path: string): Promise<void> {
  const database = await databasePromise
  await database.delete('recents', path)
  announceChange()
}

/** Empties the list, as File > Open Recent > Clear Menu does. */
export async function clearRecents(): Promise<void> {
  const database = await databasePromise
  await database.clear('recents')
  announceChange()
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

/** Calls `listener` when the list changes, in this window or another that shares the index. */
export function watchRecents(listener: () => void): () => void {
  const handleMessage = (): void => listener()
  recentEvents.addEventListener('message', handleMessage)
  localListeners.add(listener)
  return () => {
    recentEvents.removeEventListener('message', handleMessage)
    localListeners.delete(listener)
  }
}
