import { AnnotationMode, getDocument, GlobalWorkerOptions, type PDFDocumentLoadingTask, type PDFDocumentProxy, type RenderTask } from 'pdfjs-dist'
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url'

GlobalWorkerOptions.workerSrc = workerUrl

export type { PDFDocumentProxy }

const tasks = new WeakMap<Uint8Array, PDFDocumentLoadingTask>()

/** One pdf.js document per source; pdf.js takes ownership of the buffer it is given, so it gets a copy. */
export function openPdfJs(bytes: Uint8Array): Promise<PDFDocumentProxy> {
  const cached = tasks.get(bytes)
  if (cached !== undefined) return cached.promise
  const task = getDocument({ data: bytes.slice() })
  tasks.set(bytes, task)
  return task.promise
}

/** Drops a cached pdf.js document and its worker memory. */
export async function closePdfJs(bytes: Uint8Array): Promise<void> {
  const task = tasks.get(bytes)
  tasks.delete(bytes)
  await task?.destroy()
}

export interface RenderOptions {
  /** Pixels per point, before the device pixel ratio. */
  readonly scale: number
  /** Added clockwise to the page's own /Rotate. */
  readonly extraRotation: number
  /**
   * `interactive` leaves form widgets for the HTML form layer to draw; `print` bakes them in,
   * which is what thumbnails, previews and redaction rasters need.
   */
  readonly forms: 'interactive' | 'print'
  readonly pixelRatio: number
}

/** Renders one page into a canvas sized for the device; returns the running task so callers can cancel it. */
export async function renderPage(
  document: PDFDocumentProxy,
  index: number,
  canvas: HTMLCanvasElement,
  options: RenderOptions,
): Promise<RenderTask> {
  const page = await document.getPage(index + 1)
  const viewport = page.getViewport({
    scale: options.scale * options.pixelRatio,
    rotation: (page.rotate + options.extraRotation) % 360,
  })
  canvas.width = Math.max(1, Math.floor(viewport.width))
  canvas.height = Math.max(1, Math.floor(viewport.height))
  return page.render({
    canvas,
    viewport,
    // `display` renders pace on requestAnimationFrame, which Chromium stops in hidden or covered
    // windows; rasters for output (redaction, previews) use `print` so they always finish.
    intent: options.forms === 'print' ? 'print' : 'display',
    annotationMode: options.forms === 'interactive' ? AnnotationMode.ENABLE_FORMS : AnnotationMode.ENABLE,
  })
}

export function isCancelledRender(error: unknown): boolean {
  return error instanceof Error && error.name === 'RenderingCancelledException'
}
