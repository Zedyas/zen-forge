import { describe, expect, it } from 'vitest'
import { addMarkup, removeMarkup } from './markup-actions'
import { readyPdf, setPdfDocument, undoPdf } from './pdf-store'

describe('removeMarkup', () => {
  it('records nothing for a markup an undo already took away, so redo still works', () => {
    const snapshot = { pages: [{ key: 'page', source: 0, index: 0, rotation: 0 as const, markups: [] }], formValues: {}, flattenForm: false }
    setPdfDocument('doc', {
      status: 'ready', sources: [new Uint8Array([1])], sizes: new Map(), fields: [], past: [], present: snapshot, future: [], saved: snapshot,
      zoom: 1, fitted: false, currentPage: 0,
    })
    // Adding a text box selects it; undo removes it, and Delete then targets the selection left behind.
    const id = addMarkup('doc', 'page', { kind: 'text', x: 10, y: 10, text: 'Hello', size: 12, color: '#000000' })
    undoPdf('doc')

    removeMarkup('doc', 'page', id)

    expect(readyPdf('doc')?.present).toBe(snapshot)
    expect(readyPdf('doc')?.future).toHaveLength(1)
  })
})
