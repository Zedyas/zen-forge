import { useEffect, useMemo, useRef, useState } from 'react'
import type { FormField } from './engine'
import { FormLayer } from './FormLayer'
import { MarkupLayer } from './MarkupLayer'
import { PageCanvas } from './PageCanvas'
import { pageSize, sizeKey, updatePdf, type ReadyPdf } from './pdf-store'
import { select } from './tool-store'

/** The zoom at which the widest page fills the desk, kept within a comfortable reading range. */
export function fitWidthZoom(document: ReadyPdf): number | undefined {
  const desk = window.document.querySelector<HTMLElement>('.pdf-desk')
  if (desk === null || desk.clientWidth === 0) return undefined
  const widest = Math.max(...document.present.pages.map(item => pageSize(document, item).width))
  return Math.min(1.5, Math.max(0.5, Math.floor(((desk.clientWidth - 64) / widest) * 20) / 20))
}

export function scrollToPage(pageKey: string): void {
  window.document.querySelector(`.pdf-desk [data-page-key="${pageKey}"]`)?.scrollIntoView({ block: 'start' })
}

interface PdfDeskProps {
  readonly documentId: string
  readonly document: ReadyPdf
}

/** Continuous vertical pages; tracks which page is current as the desk scrolls. */
export function PdfDesk({ documentId, document }: PdfDeskProps) {
  const [desk, setDesk] = useState<HTMLDivElement | null>(null)
  const frame = useRef(0)
  const { zoom, present } = document

  useEffect(() => {
    if (desk === null || document.fitted) return
    const zoomToFit = fitWidthZoom(document)
    if (zoomToFit !== undefined) updatePdf(documentId, () => ({ zoom: Math.min(zoomToFit, 1.25), fitted: true }))
  }, [desk, document, documentId])

  useEffect(() => () => cancelAnimationFrame(frame.current), [])

  const fieldsByPage = useMemo(() => {
    const byPage = new Map<number, FormField[]>()
    document.fields.forEach(field => field.widgets.forEach(widget => {
      const list = byPage.get(widget.page) ?? []
      if (!list.includes(field)) list.push(field)
      byPage.set(widget.page, list)
    }))
    return byPage
  }, [document.fields])

  const trackCurrentPage = (): void => {
    cancelAnimationFrame(frame.current)
    frame.current = requestAnimationFrame(() => {
      if (desk === null) return
      const line = desk.scrollTop + desk.clientHeight * 0.3
      const pages = Array.from(desk.querySelectorAll<HTMLElement>('[data-page-key]'))
      const index = pages.findIndex(page => page.offsetTop + page.offsetHeight > line)
      const current = index < 0 ? pages.length - 1 : index
      if (current !== document.currentPage) updatePdf(documentId, () => ({ currentPage: current }))
    })
  }

  return (
    <div ref={setDesk} className="pdf-desk" onScroll={trackCurrentPage}>
      <div className="pdf-pages">
        {present.pages.map(item => {
          const size = pageSize(document, item)
          const source = document.sources[item.source]
          // Form inputs only sit on the original document's pages; pages from inserted files carry none.
          const fields = item.source === 0 ? fieldsByPage.get(item.index) : undefined
          if (source === undefined) return null
          return (
            <div
              key={item.key}
              data-page-key={item.key}
              className="pdf-page"
              style={{ width: size.width * zoom, height: size.height * zoom }}
              onPointerDown={event => {
                if (!(event.target instanceof Element) || event.target.closest('.pdf-markup') === null) select(undefined)
              }}
            >
              <PageCanvas
                source={source}
                index={item.index}
                extraRotation={item.rotation}
                scale={zoom}
                forms={fields === undefined ? 'print' : 'interactive'}
                root={desk}
                margin={900}
              />
              {fields !== undefined && (
                <FormLayer
                  documentId={documentId}
                  fields={fields}
                  pageIndex={item.index}
                  rotation={item.rotation}
                  unrotatedSize={document.sizes.get(sizeKey(item.source, item.index)) ?? size}
                  values={present.formValues}
                  zoom={zoom}
                />
              )}
              <MarkupLayer documentId={documentId} item={item} size={size} zoom={zoom} />
            </div>
          )
        })}
      </div>
    </div>
  )
}
