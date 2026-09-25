import { beforeEach, describe, expect, it } from 'vitest'
import type { FileReference } from '@shared/shell'
import { homeTabId, useDocumentsStore } from './documents-store'

function file(path: string): FileReference {
  const name = path.split('/').pop() ?? path
  return { path, name, extension: name.split('.').pop() ?? '', size: 1, lastModified: 0 }
}

function tabIds(): readonly string[] {
  return useDocumentsStore.getState().documents.map(document => document.id)
}

describe('documents store', () => {
  beforeEach(() => useDocumentsStore.setState({ documents: [], activeId: homeTabId }))

  it('keeps focus on a file opened while last session is being restored', () => {
    const store = useDocumentsStore.getState()
    store.openFile(file('/from-finder.pdf'))
    store.restoreFile(file('/last-session.pdf'))

    expect(useDocumentsStore.getState().activeId).toBe('/from-finder.pdf')
    expect(tabIds()).toEqual(['/from-finder.pdf', '/last-session.pdf'])
  })

  it('opens a new tab at the end of its own kind, keeping spreadsheets and PDFs grouped', () => {
    const store = useDocumentsStore.getState()
    store.openFile(file('/budget.xlsx'))
    store.openFile(file('/lease.pdf'))
    store.openFile(file('/sales.csv'))
    store.openUntitled('sheets')
    store.openFile(file('/invoice.pdf'))

    expect(tabIds()).toEqual(['/budget.xlsx', '/sales.csv', expect.stringMatching(/^untitled:/), '/lease.pdf', '/invoice.pdf'])
  })

  it('selects the right-hand neighbour when the active tab closes, and Home after the last one', () => {
    const store = useDocumentsStore.getState()
    store.openFile(file('/a.pdf'))
    store.openFile(file('/b.pdf'))
    store.select('/a.pdf')
    store.remove('/a.pdf')
    expect(useDocumentsStore.getState().activeId).toBe('/b.pdf')

    store.remove('/b.pdf')
    expect(useDocumentsStore.getState().activeId).toBe(homeTabId)
  })

  it('steps through tabs with Home as the first, wrapping at both ends', () => {
    const store = useDocumentsStore.getState()
    store.openFile(file('/a.pdf'))
    store.openFile(file('/b.pdf'))

    store.selectRelative(1)
    expect(useDocumentsStore.getState().activeId).toBe(homeTabId)
    store.selectRelative(-1)
    expect(useDocumentsStore.getState().activeId).toBe('/b.pdf')
    store.selectIndex(Number.MAX_SAFE_INTEGER)
    expect(useDocumentsStore.getState().activeId).toBe('/b.pdf')
    store.selectIndex(1)
    expect(useDocumentsStore.getState().activeId).toBe('/a.pdf')
  })
})
