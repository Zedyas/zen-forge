import { flushSync } from 'react-dom'
import { createRoot } from 'react-dom/client'
import { toast } from 'sonner'
import { openFiles } from '../../app/document-actions'
import { useDocumentsStore } from '../../app/documents-store'
import { fileService } from '../../services/file/IpcFileService'
import { openPrintDialog, printToPdf, type Printout } from '../../services/print/print'
import type { Presentation } from './model'
import { finishTyping } from './slides-actions'
import { readySlides } from './slides-store'
import { SlideView } from './SlideView'

/*
 * A printout is the static SlideView of every slide, one per sheet of paper at the slide's own
 * size, so a printed or exported slide looks exactly as it does on screen. Speaker notes are not
 * printed.
 */

const millimetresPerPoint = 25.4 / 72
const cssPixelsPerPoint = 96 / 72

function slidesPrintout(presentation: Presentation): Printout {
  const { width, height } = presentation
  const page = { width: width * millimetresPerPoint, height: height * millimetresPerPoint, margin: 0 }
  // A fraction of a millimetre short of the page, so rounding never spills a slide onto a second sheet.
  const scale = cssPixelsPerPoint * ((page.height - 0.3) / page.height)
  const content = document.createElement('div')
  content.className = 'slides-print'
  const root = createRoot(content)
  flushSync(() => root.render(presentation.slides.map(slide => (
    <div key={slide.id} className="slides-print-page">
      <SlideView slide={slide} width={width} height={height} scale={scale} />
    </div>
  ))))
  return { content, page, release: () => root.unmount() }
}

export async function printPresentation(id: string): Promise<void> {
  finishTyping(id)
  const document = readySlides(id)
  if (document !== undefined) await openPrintDialog(() => slidesPrintout(document.present))
}

/** Writes every slide, as it prints, to a PDF chosen by the user, and offers to open it in Hanko. */
export async function exportPresentationPdf(id: string): Promise<void> {
  finishTyping(id)
  const document = readySlides(id)
  const open = useDocumentsStore.getState().documents.find(candidate => candidate.id === id)
  if (document === undefined || open === undefined) return
  const path = await fileService.chooseSavePath({ defaultName: `${open.name}.pdf`, extensions: ['pdf'] })
  if (path === undefined) return
  const bytes = await printToPdf(() => slidesPrintout(document.present))
  if (bytes === undefined) return
  await fileService.write(path, bytes)
  const file = await fileService.describe(path)
  toast.success('Exported as PDF', {
    description: file.name,
    action: { label: 'Open', onClick: () => void openFiles([file]) },
  })
}
