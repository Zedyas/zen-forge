import { useEffect, useRef, useState } from 'react'
import type { RenderTask } from 'pdfjs-dist'
import { isCancelledRender, openPdfJs, renderPage } from './pdfjs'

interface PageCanvasProps {
  readonly source: Uint8Array
  readonly index: number
  readonly extraRotation: number
  /** CSS pixels per point. */
  readonly scale: number
  readonly forms: 'interactive' | 'print'
  /** The scrolling element the page lives in; rendering starts `margin` pixels before it scrolls into view. */
  readonly root: Element | null
  readonly margin: number
}

/**
 * Renders one page only while it is near the viewport, and releases the bitmap when it leaves,
 * so a long document holds a few pages of pixels rather than all of them.
 */
export function PageCanvas({ source, index, extraRotation, scale, forms, root, margin }: PageCanvasProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const [visible, setVisible] = useState(false)
  const [failed, setFailed] = useState(false)

  useEffect(() => {
    const canvas = canvasRef.current
    if (canvas === null || root === null) return
    const observer = new IntersectionObserver(
      entries => setVisible(entries.some(entry => entry.isIntersecting)),
      { root, rootMargin: `${margin}px 0px` },
    )
    observer.observe(canvas)
    return () => observer.disconnect()
  }, [root, margin])

  useEffect(() => {
    const canvas = canvasRef.current
    if (canvas === null) return
    if (!visible) {
      canvas.width = 0
      canvas.height = 0
      return
    }
    let task: RenderTask | undefined
    let cancelled = false
    void (async () => {
      try {
        const document = await openPdfJs(source)
        if (cancelled) return
        task = await renderPage(document, index, canvas, { scale, extraRotation, forms, pixelRatio: window.devicePixelRatio })
        if (cancelled) task.cancel()
        await task.promise
        setFailed(false)
      } catch (error) {
        if (!cancelled && !isCancelledRender(error)) setFailed(true)
      }
    })()
    return () => {
      cancelled = true
      task?.cancel()
    }
  }, [visible, source, index, extraRotation, scale, forms])

  return (
    <>
      <canvas ref={canvasRef} className="pdf-canvas" />
      {failed && <p className="pdf-page-error">This page could not be drawn.</p>}
    </>
  )
}
