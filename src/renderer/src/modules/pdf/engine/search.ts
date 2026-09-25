import { rectToDisplayed, type PageGeometry, type Rect } from './geometry'

/** A page to search, and how the editor shows it: the rectangles found are placed on that. */
export interface SearchPage {
  readonly index: number
  readonly shown: PageGeometry
}

/**
 * Where `query` appears on some pages of one PDF, as MuPDF reads their text: for each page, one
 * list of rectangles per match (a match can span lines), in displayed page coordinates. Case is
 * ignored, and a space in the query matches any whitespace, line breaks included, as in find. MuPDF places every glyph, so these rectangles are exact even on justified lines,
 * where pdf.js knows only a run's total width and its word and letter spacing is spread evenly.
 * MuPDF returns at most 500 rectangles a page; any matches past those are missing.
 */
export async function searchText(bytes: Uint8Array, pages: readonly SearchPage[], query: string): Promise<Rect[][][]> {
  const needle = query.trim().split(/\s+/).join(' ')
  if (needle === '') return pages.map(() => [])
  // Loaded on first use, as for redaction.
  const mupdf = await import('mupdf')
  const document = new mupdf.PDFDocument(bytes)
  try {
    return pages.map(({ index, shown }) => {
      const page = document.loadPage(index)
      try {
        // Results are in MuPDF's page space, which applies the page's rotation, crop box and UserUnit
        // in its own way. Its transform takes user space there, so the inverse leads back to user
        // space, and the editor's own geometry leads on to the page as shown.
        const [a, b, c, d, e, f] = mupdf.Matrix.invert(page.getTransform())
        return page.search(needle, 'ignore-case').map(match => match.map(quad => {
          const xs: number[] = []
          const ys: number[] = []
          for (let corner = 0; corner < 8; corner += 2) {
            const x = quad[corner] ?? 0
            const y = quad[corner + 1] ?? 0
            xs.push(a * x + c * y + e)
            ys.push(b * x + d * y + f)
          }
          const user = { x: Math.min(...xs), y: Math.min(...ys), width: Math.max(...xs) - Math.min(...xs), height: Math.max(...ys) - Math.min(...ys) }
          return rectToDisplayed(shown, user)
        }))
      } finally {
        page.destroy()
      }
    })
  } finally {
    document.destroy()
  }
}
