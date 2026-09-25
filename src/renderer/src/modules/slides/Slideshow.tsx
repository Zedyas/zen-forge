import { useEffect, useState } from 'react'
import { createPortal } from 'react-dom'
import { stepSlideshow, stopSlideshow } from './slides-actions'
import type { ReadySlides } from './slides-store'
import { SlideView } from './SlideView'

const nextKeys: ReadonlySet<string> = new Set(['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter', 'n'])
const previousKeys: ReadonlySet<string> = new Set(['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p'])

function useWindowSize(): { readonly width: number; readonly height: number } {
  const [size, setSize] = useState({ width: window.innerWidth, height: window.innerHeight })
  useEffect(() => {
    const resize = (): void => setSize({ width: window.innerWidth, height: window.innerHeight })
    window.addEventListener('resize', resize)
    return () => window.removeEventListener('resize', resize)
  }, [])
  return size
}

/** Plays the presentation over the whole window: one slide at a time, scaled to fit on black. */
export function Slideshow({ documentId, document }: { readonly documentId: string; readonly document: ReadySlides }) {
  const size = useWindowSize()
  const index = document.playing ?? 0
  const slide = document.present.slides[index]
  const last = document.present.slides.length - 1

  useEffect(() => {
    // Capture phase, so the editor's own keys (nudging, deleting) never see the slideshow's.
    const handleKeyDown = (event: KeyboardEvent): void => {
      if (event.metaKey || event.ctrlKey || event.altKey) return
      let to: number | undefined
      if (event.key === 'Escape') {
        event.preventDefault()
        event.stopPropagation()
        stopSlideshow(documentId)
        return
      }
      if (nextKeys.has(event.key)) to = index + 1
      else if (previousKeys.has(event.key)) to = index - 1
      else if (event.key === 'Home') to = 0
      else if (event.key === 'End') to = last
      if (to === undefined) return
      event.preventDefault()
      event.stopPropagation()
      stepSlideshow(documentId, to)
    }
    window.addEventListener('keydown', handleKeyDown, true)
    return () => window.removeEventListener('keydown', handleKeyDown, true)
  }, [documentId, index, last])

  // Closing or leaving the tab removes the show without Escape; the window leaves full screen too.
  useEffect(() => () => stopSlideshow(documentId), [documentId])

  if (slide === undefined) return null
  const { width, height } = document.present
  const scale = Math.min(size.width / width, size.height / height)

  return createPortal(
    <div
      className="slideshow"
      role="dialog"
      aria-label={`Slide ${index + 1} of ${last + 1}`}
      onClick={() => stepSlideshow(documentId, index + 1)}
      onContextMenu={event => {
        event.preventDefault()
        stepSlideshow(documentId, index - 1)
      }}
    >
      <SlideView slide={slide} width={width} height={height} scale={scale} />
    </div>,
    window.document.body,
  )
}
