import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { OpenDocument } from '../../app/documents-store'
import type { EditorHandler } from '../../app/editors'
import { useViewStore } from '../../app/view-store'
import { FidelitySurface } from '../../ui/FidelitySurface'
import { WindowEmpty } from '../../ui/WindowEmpty'
import { ConfirmHost } from '../pdf/ConfirmDialog'
import { holdsText } from './model'
import {
  deleteSlide,
  duplicateSelection,
  isTextEntry,
  loadSlidesDocument,
  nudgeSelection,
  releaseSlidesDocument,
  removeElement,
  runSlidesCommand,
  saveSlidesDocument,
  select,
  setSlideNotes,
  showSlide,
  startEditing,
  startNewPresentation,
} from './slides-actions'
import { currentSlide, readySlides, selectedElement, useSlidesStore, type ReadySlides } from './slides-store'
import { SlideCanvas } from './SlideCanvas'
import { SlideRail } from './SlideRail'
import { SlidesInspector } from './SlidesInspector'
import { SlidesToolbar } from './SlidesToolbar'
import { activeEditor } from './text-editing'
import './slides.css'

/**
 * Keys on the canvas: Delete, arrow nudges (Shift for 10 pt), ⌘D, Return to type into the selected
 * box, Escape to deselect. With the slide rail focused, arrows change slide and Delete removes it.
 * Ignored while typing.
 */
function useSlideKeys(documentId: string | undefined): void {
  useEffect(() => {
    if (documentId === undefined) return
    const handleKeyDown = (event: KeyboardEvent): void => {
      const document = readySlides(documentId)
      if (document === undefined || isTextEntry(event.target)) return
      const inRail = event.target instanceof Element && event.target.closest('.slides-rail') !== null
      const element = selectedElement(document)
      if (event.metaKey && !event.shiftKey && !event.altKey && event.key.toLowerCase() === 'd') {
        event.preventDefault()
        duplicateSelection(documentId)
        return
      }
      if (event.metaKey || event.ctrlKey || event.altKey) return
      const slides = document.present.slides
      const index = slides.findIndex(slide => slide.id === document.slideId)
      if (inRail || element === undefined) {
        const step = event.key === 'ArrowUp' || event.key === 'PageUp' ? -1 : event.key === 'ArrowDown' || event.key === 'PageDown' ? 1 : 0
        const target = slides[index + step]
        if (step !== 0 && target !== undefined) {
          event.preventDefault()
          showSlide(documentId, target.id)
          return
        }
        if (inRail && (event.key === 'Backspace' || event.key === 'Delete')) {
          event.preventDefault()
          deleteSlide(documentId)
        }
        return
      }
      if (event.key === 'Escape') {
        select(documentId, undefined)
        return
      }
      if (event.key === 'Backspace' || event.key === 'Delete') {
        event.preventDefault()
        removeElement(documentId, element.id)
        return
      }
      if (event.key === 'Enter' && holdsText(element)) {
        event.preventDefault()
        startEditing(documentId, element.id)
        return
      }
      const nudge = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[event.key]
      if (nudge !== undefined) {
        event.preventDefault()
        const step = event.shiftKey ? 10 : 1
        nudgeSelection(documentId, (nudge[0] ?? 0) * step, (nudge[1] ?? 0) * step)
      }
    }
    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [documentId])
}

/** Speaker notes for one slide; written to the presentation when the field loses focus or goes away. */
function NotesField({ documentId, slideId, notes }: { readonly documentId: string; readonly slideId: string; readonly notes: string }) {
  const [draft, setDraft] = useState<string>()
  const typed = useRef<string | undefined>(undefined)
  const commit = (): void => {
    if (typed.current !== undefined) setSlideNotes(documentId, slideId, typed.current)
    typed.current = undefined
  }
  // Switching tabs removes the field without a blur; what was typed is kept.
  useLayoutEffect(() => () => {
    if (typed.current !== undefined) setSlideNotes(documentId, slideId, typed.current)
  }, [documentId, slideId])
  return (
    <textarea
      className="slides-notes"
      value={draft ?? notes}
      placeholder="Speaker notes"
      aria-label="Speaker notes"
      spellCheck
      onFocus={() => activeEditor(documentId)?.commit()}
      onChange={event => {
        typed.current = event.currentTarget.value
        setDraft(event.currentTarget.value)
      }}
      onBlur={() => {
        commit()
        setDraft(undefined)
      }}
    />
  )
}

function StatusBar({ documentId, document }: { readonly documentId: string; readonly document: ReadySlides }) {
  const index = document.present.slides.findIndex(slide => slide.id === document.slideId)
  return (
    <footer className="statusbar">
      <span className="statusbar-item">Slide <b>{index + 1}</b> of {document.present.slides.length}</span>
      <span className="statusbar-grow" />
      <FidelitySurface documentId={documentId} />
    </footer>
  )
}

export const slidesEditor: EditorHandler = { run: runSlidesCommand, save: saveSlidesDocument, release: releaseSlidesDocument }

/** Slides: the presentation tab's content below the title bar. */
export function SlidesWorkspace({ document }: { readonly document: OpenDocument }) {
  const documentId = document.id
  const state = useSlidesStore(store => store.documents[documentId])
  const inspector = useViewStore(view => view.inspector)
  useSlideKeys(state?.status === 'ready' ? documentId : undefined)

  useEffect(() => {
    if (useSlidesStore.getState().documents[document.id] !== undefined) return
    if (document.path === undefined) startNewPresentation(document.id)
    else void loadSlidesDocument(document.id, document.path, `${document.name}.${document.extension}`)
  }, [document])

  if (state === undefined || state.status === 'loading') {
    return <section className="window-empty"><p>Opening {document.name}…</p></section>
  }

  if (state.status === 'error') {
    return <WindowEmpty application="slides" title={`${document.name} could not be opened`} description={state.message} />
  }

  const slide = currentSlide(state)
  return (
    <>
      <ConfirmHost />
      <SlidesToolbar documentId={documentId} document={state} />
      <div className="window-body">
        <SlideRail documentId={documentId} document={state} />
        <div className="window-main">
          <div className="slides-main">
            <SlideCanvas documentId={documentId} document={state} />
            {slide !== undefined && <NotesField key={slide.id} documentId={documentId} slideId={slide.id} notes={slide.notes} />}
          </div>
          <StatusBar documentId={documentId} document={state} />
        </div>
        {inspector && <SlidesInspector documentId={documentId} document={state} openDocument={document} />}
      </div>
    </>
  )
}
