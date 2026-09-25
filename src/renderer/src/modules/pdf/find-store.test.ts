import { describe, expect, it, vi } from 'vitest'
import { findInDocument } from './find-store'
import { releasePdf, setPdfDocument } from './pdf-store'
import { openPdfJs, type PDFDocumentProxy } from './pdfjs'

// Once a PDF has closed, each call would open a pdf.js document, with its own worker, that nothing closes.
vi.mock('./pdfjs', () => ({ openPdfJs: vi.fn() }))

describe('findInDocument', () => {
  it('loads no more pages once the PDF closes', async () => {
    const snapshot = { pages: [0, 1, 2].map(index => ({ key: `page-${index}`, source: 0, index, rotation: 0 as const, markups: [] })), formValues: {}, flattenForm: false }
    setPdfDocument('doc', {
      status: 'ready', sources: [new Uint8Array([1])], sizes: new Map(), fields: [], past: [], present: snapshot, future: [], saved: snapshot,
      zoom: 1, fitted: false, currentPage: 0,
    })
    const page = {
      rotate: 0,
      getViewport: () => ({ transform: [1, 0, 0, -1, 0, 792] }),
      // The tab closes while the first page's text loads.
      getTextContent: async () => {
        releasePdf('doc')
        return { items: [], styles: {} }
      },
    }
    vi.mocked(openPdfJs).mockImplementation(async () => ({ getPage: async () => page }) as unknown as PDFDocumentProxy)

    await findInDocument('doc', 'rent', true)
    expect(openPdfJs).toHaveBeenCalledTimes(1)
  })
})
