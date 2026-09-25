import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { OpenDocument } from '../../app/documents-store'
import type { EditorHandler } from '../../app/editors'
import { useViewStore } from '../../app/view-store'
import { FidelitySurface } from '../../ui/FidelitySurface'
import { WindowEmpty } from '../../ui/WindowEmpty'
import { ConfirmHost } from '../pdf/ConfirmDialog'
import { holdsText } from './model'
import { runSlidesCommand } from './slide-commands'
import {
  copySelection,
  deleteSlide,
  duplicateSelection,
  finishTyping,
  isTextEntry,
  loadSlidesDocument,
  nudgeSelection,
  paste,
  releaseSlidesDocument,
  removeSelection,
  saveSlidesDocument,
  select,
  setSlideNotes,
  showSlide,
  startEditing,
  startNewPresentation,
} from './slides-actions'
import { currentSlide, markTyping, readySlides, selectedElement, typingDone, useSlidesStore, type ReadySlides } from './slides-store'
import { SlideCanvas } from './SlideCanvas'
import { SlidesErrorBoundary } from './SlidesErrorBoundary'
import { SlideRail } from './SlideRail'
import { SlidesInspector } from './SlidesInspector'
import { Slideshow } from './Slideshow'
import { SlidesToolbar } from './SlidesToolbar'
import { activeEditor } from './text-editing'
import './slides.css'

/**
 * Keys on the canvas: Delete, arrow nudges (Shift for 10 pt), ⌘D, Return to type into the selected
 * box, Escape to deselect. With the slide rail focused, or nothing selected, arrows change slide
 * and, in the rail, Delete removes it. Ignored while typing and while the slideshow plays.
 */
function useSlideKeys(documentId: string | undefined): void {
  useEffect(() => {
    if (documentId === undefined) return
    const handleKeyDown = (event: KeyboardEvent): void => {
      const document = readySlides(documentId)
      if (document === undefined || document.playing !== undefined || event.defaultPrevented || isTextEntry(event.target)) return
      // Return and Space press a focused button or menu item rather than editing the selection.
      const control = event.target instanceof Element && event.target.closest('button, [role="menuitem"], [role="menu"], select, a') !== null
      if (control && (event.key === 'Enter' || event.key === ' ')) return
      const inRail = event.target instanceof Element && event.target.closest('.slides-rail') !== null
      if (event.metaKey && !event.shiftKey && !event.altKey && event.key.toLowerCase() === 'd') {
        event.preventDefault()
        duplicateSelection(documentId)
        return
      }
      if (event.metaKey || event.ctrlKey || event.altKey) return
      if (inRail || document.selection.length === 0) {
        const slides = document.present.slides
        const index = slides.findIndex(slide => slide.id === document.slideId)
        const step = event.key === 'ArrowUp' || event.key === 'PageUp' ? -1 : event.key === 'ArrowDown' || event.key === 'PageDown' ? 1 : 0
        const target = slides[index + step]
        if (step !== 0 && target !== undefined) {
          event.preventDefault()
          showSlide(documentId, target.id)
        } else if (inRail && (event.key === 'Backspace' || event.key === 'Delete')) {
          event.preventDefault()
          deleteSlide(documentId)
        }
        return
      }
      const single = selectedElement(document)
      if (event.key === 'Escape') {
        select(documentId, [])
      } else if (event.key === 'Backspace' || event.key === 'Delete') {
        event.preventDefault()
        removeSelection(documentId)
      } else if (event.key === 'Enter' && single !== undefined && (holdsText(single) || single.kind === 'table')) {
        event.preventDefault()
        startEditing(documentId, single.id, single.kind === 'table' ? document.cell ?? { row: 0, column: 0 } : undefined)
      } else {
        const nudge = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[event.key]
        if (nudge === undefined) return
        event.preventDefault()
        const step = event.shiftKey ? 10 : 1
        nudgeSelection(documentId, (nudge[0] ?? 0) * step, (nudge[1] ?? 0) * step)
      }
    }
    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [documentId])
}

/**
 * Edit > Copy, Cut and Paste reach the page as clipboard events. Outside a text field they copy the
 * selected objects, and paste objects, a picture or text onto the current slide.
 */
function useSlideClipboard(documentId: string | undefined): void {
  useEffect(() => {
    if (documentId === undefined) return
    const handle = (kind: 'copy' | 'cut' | 'paste') => (event: ClipboardEvent): void => {
      const document = readySlides(documentId)
      if (document === undefined || document.playing !== undefined || isTextEntry(event.target) || event.clipboardData === null) return
      const handled = kind === 'paste' ? paste(documentId, event.clipboardData) : copySelection(documentId, event.clipboardData, kind === 'cut')
      if (handled) event.preventDefault()
    }
    const copy = handle('copy')
    const cut = handle('cut')
    const pasted = handle('paste')
    window.document.addEventListener('copy', copy)
    window.document.addEventListener('cut', cut)
    window.document.addEventListener('paste', pasted)
    return () => {
      window.document.removeEventListener('copy', copy)
      window.document.removeEventListener('cut', cut)
      window.document.removeEventListener('paste', pasted)
    }
  }, [documentId])
}

/** Speaker notes for one slide; written to the presentation when the field loses focus or goes away. */
function NotesField({ documentId, slideId, notes }: { readonly documentId: string; readonly slideId: string; readonly notes: string }) {
  const [draft, setDraft] = useState<string>()
  const typed = useRef<string | undefined>(undefined)
  const commit = (): void => {
    if (typed.current !== undefined) setSlideNotes(documentId, slideId, typed.current)
    typed.current = undefined
    typingDone(documentId)
  }
  // Switching tabs removes the field without a blur; what was typed is kept.
  useLayoutEffect(() => () => {
    if (typed.current === undefined) return
    setSlideNotes(documentId, slideId, typed.current)
    typingDone(documentId)
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
        if (typed.current === undefined) markTyping(documentId)
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

export const slidesEditor: EditorHandler = {
  run: runSlidesCommand,
  save: saveSlidesDocument,
  release: releaseSlidesDocument,
  flush: finishTyping,
  presenting: id => readySlides(id)?.playing !== undefined,
}

/** Slides: the presentation tab's content below the title bar. */
export function SlidesWorkspace({ document }: { readonly document: OpenDocument }) {
  const documentId = document.id
  const state = useSlidesStore(store => store.documents[documentId])
  const inspector = useViewStore(view => view.inspector)
  useSlideKeys(state?.status === 'ready' ? documentId : undefined)
  useSlideClipboard(state?.status === 'ready' ? documentId : undefined)

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
    <SlidesErrorBoundary key={documentId} documentId={documentId}>
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
      {state.playing !== undefined && <Slideshow documentId={documentId} document={state} />}
    </SlidesErrorBoundary>
  )
}
