import { toast } from 'sonner'
import type { Paper } from '@shared/print'
import { openPrintDialog, pageSetup, printableArea, type Orientation, type PageSetup, type Printout } from '../../services/print/print'
import { sameGeometry, savePdf } from './engine'
import { isRedaction, toPageRef, type RedactBox } from './model'
import { closePdfJs, openPdfJs, renderPage, shownPage, type PDFDocumentProxy } from './pdfjs'
import { readyPdf, type ReadyPdf } from './pdf-store'

/*
 * Printing prints each page as an image, drawn from the document as it is now, unsaved changes
 * included. This is deliberate: redaction boxes are painted black onto the pixels, so the text
 * under them is in no layer of the printout. The print dialog can Save as PDF, and printing the
 * pages' real text would put redacted text into that file. Save runs MuPDF's redaction instead,
 * which keeps the rest of the text; printing does not need it.
 */

/** Long side of each page image in pixels: about 250 dpi on Letter or A4, sharp on paper without huge memory use. */
const rasterLongSide = 2800
/** Blank space around each page, in millimetres, so printers that cannot reach the edge keep all of it. */
const margin = 6
/** Longer documents show their progress while the pages are drawn. */
const quietPageCount = 3

interface PageImage {
  readonly url: string
  /** The page's displayed size, in points. */
  readonly width: number
  readonly height: number
}

async function drawPage(pdf: PDFDocumentProxy, index: number, boxes: readonly RedactBox[]): Promise<PageImage> {
  const page = await pdf.getPage(index + 1)
  const { width, height } = page.getViewport({ scale: 1 })
  const canvas = document.createElement('canvas')
  try {
    const task = await renderPage(pdf, index, canvas, {
      scale: rasterLongSide / Math.max(width, height),
      extraRotation: 0,
      forms: 'print',
      pixelRatio: 1,
    })
    await task.promise
    const context = canvas.getContext('2d')
    if (context === null) throw new Error('A page could not be drawn for printing.')
    const scaleX = canvas.width / width
    const scaleY = canvas.height / height
    context.fillStyle = '#000'
    // Rounded outwards, so no anti-aliased edge of covered text shows beside a box.
    for (const box of boxes) {
      const left = Math.floor(box.x * scaleX)
      const top = Math.floor(box.y * scaleY)
      context.fillRect(left, top, Math.ceil((box.x + box.width) * scaleX) - left, Math.ceil((box.y + box.height) * scaleY) - top)
    }
    const image = await new Promise<Blob | null>(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.92))
    if (image === null) throw new Error('A page could not be drawn for printing.')
    return { url: URL.createObjectURL(image), width, height }
  } finally {
    // A canvas keeps its pixels until it is resized; free them now rather than at garbage collection.
    canvas.width = 0
    canvas.height = 0
    page.cleanup()
  }
}

/**
 * Draws every page from the bytes a save would write, before redaction, then paints the redaction
 * boxes. Resolves undefined when the document closes meanwhile, so nothing is printed for it.
 */
async function drawPages(id: string, state: ReadyPdf): Promise<PageImage[] | undefined> {
  const { pages, formValues, flattenForm } = state.present
  const progress = pages.length > quietPageCount ? toast('Preparing to print…', { duration: Infinity }) : undefined
  const images: PageImage[] = []
  let bytes: Uint8Array | undefined
  let drawn = false
  try {
    bytes = await savePdf({ sources: state.sources, pages: pages.map(toPageRef), formValues, flattenForm })
    const pdf = await openPdfJs(bytes)
    for (const [index, item] of pages.entries()) {
      // A closed document's pdf.js copies are gone, and reading its pages would open new ones that nothing closes.
      if (readyPdf(id) === undefined) return undefined
      if (progress !== undefined) toast(`Preparing page ${index + 1} of ${pages.length}…`, { id: progress, duration: Infinity })
      const boxes = item.markups.flatMap(placed => isRedaction(placed.markup) ? [placed.markup] : [])
      // Redaction boxes are painted where they were drawn on the page shown, so a page with boxes must print as shown.
      if (boxes.length > 0 && !sameGeometry(await shownPage(bytes, index, 0), await shownPage(state.sources[item.source], item.index, item.rotation))) {
        throw new Error(`Zendo can’t print page ${index + 1} safely: the printed page would not match the page shown, so redaction boxes could miss. Print this PDF from Adobe Acrobat instead.`)
      }
      images.push(await drawPage(pdf, index, boxes))
    }
    drawn = true
    return images
  } finally {
    if (!drawn) images.forEach(image => URL.revokeObjectURL(image.url))
    if (progress !== undefined) toast.dismiss(progress)
    if (bytes !== undefined) await closePdfJs(bytes)
  }
}

/** One page image per sheet of paper, turned to the page's orientation and fitted inside the margins. Styles are in pdf.css. */
function pagesPrintout(images: readonly PageImage[], paper: Paper): Printout {
  const setups: Record<Orientation, PageSetup> = {
    portrait: pageSetup(paper, 'portrait', margin),
    landscape: pageSetup(paper, 'landscape', margin),
  }
  const content = document.createElement('div')
  content.className = 'pdf-print'
  for (const image of images) {
    const orientation: Orientation = image.width > image.height ? 'landscape' : 'portrait'
    const area = printableArea(setups[orientation])
    // A millimetre short of the full height, so rounding never pushes a sheet onto a second page.
    const height = area.height - 1
    // Millimetres of paper per point of page: the page fills the area in one direction.
    const fit = Math.min(area.width / image.width, height / image.height)
    const sheet = content.appendChild(document.createElement('div'))
    sheet.style.setProperty('page', orientation)
    sheet.style.height = `${height}mm`
    const picture = sheet.appendChild(document.createElement('img'))
    picture.src = image.url
    picture.style.width = `${image.width * fit}mm`
    picture.style.height = `${image.height * fit}mm`
  }
  return {
    content,
    page: setups.portrait,
    namedPages: setups,
    release: () => images.forEach(image => URL.revokeObjectURL(image.url)),
  }
}

/** Opens the print dialog for a PDF as it is now, unsaved changes included. */
export async function printPdfDocument(id: string): Promise<void> {
  // A text box or form field being typed in commits on blur; do that first so the printout includes it.
  if (document.activeElement instanceof HTMLElement) document.activeElement.blur()
  const state = readyPdf(id)
  if (state === undefined) return
  await openPrintDialog(async paper => {
    const images = await drawPages(id, state)
    return images === undefined ? undefined : pagesPrintout(images, paper)
  })
}
