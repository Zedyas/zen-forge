// @vitest-environment happy-dom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'

const { setFullScreen } = vi.hoisted(() => ({ setFullScreen: vi.fn<(fullScreen: boolean) => Promise<void>>(async () => undefined) }))
vi.mock('../../services/index/document-index', () => ({ recordRecent: async () => undefined }))
vi.mock('../../services/platform/client', () => ({ platformClient: { setFullScreen } }))

import { useDocumentsStore } from '../../app/documents-store'
import { editTable } from './format-actions'
import { newPresentation, newTable, updateCell, updateSlide } from './model'
import { SlideCanvas } from './SlideCanvas'
import { finishTyping, paste, releaseSlidesDocument, runSlideCommand, startEditing, startSlideshow } from './slides-actions'
import { currentSlide, readyDocument, readySlides, setSlidesDocument, updateSlides, useSlidesStore } from './slides-store'
import { Slideshow } from './Slideshow'
import { TextEditor } from './TextEditor'

Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })

const cleanups: Array<() => void> = []
afterEach(() => cleanups.splice(0).forEach(cleanup => cleanup()))

function mount(): ReturnType<typeof createRoot> {
  const host = document.createElement('div')
  document.body.append(host)
  const root = createRoot(host)
  cleanups.push(() => host.remove())
  return root
}

function isDirty(id: string): boolean | undefined {
  return useDocumentsStore.getState().documents.find(candidate => candidate.id === id)?.dirty
}

describe('typing in a text box', () => {
  it('marks the tab unsaved at the first keystroke, and clean again when the edit came to nothing', async () => {
    if (typeof document.execCommand !== 'function') document.execCommand = () => true
    const id = useDocumentsStore.getState().openUntitled('slides')
    setSlidesDocument(id, readyDocument(newPresentation()))
    const element = readySlides(id)?.present.slides[0]?.elements[0]
    if (element?.kind !== 'text') throw new Error('The new presentation has no title box')
    const root = mount()
    await act(async () => root.render(<TextEditor documentId={id} element={element} />))
    const editor = document.querySelector('[contenteditable]')
    const paragraph = editor?.querySelector('p')
    if (!(editor instanceof HTMLElement) || paragraph === null || paragraph === undefined) throw new Error('No editor')

    paragraph.textContent = 'Quarterly review'
    await act(async () => editor.dispatchEvent(new Event('input', { bubbles: true })))
    // A window close or quit reads this before anything is written into the presentation.
    expect(isDirty(id)).toBe(true)

    paragraph.innerHTML = '<br>'
    await act(async () => editor.dispatchEvent(new Event('input', { bubbles: true })))
    await act(async () => finishTyping(id))
    expect(isDirty(id)).toBe(false)
    await act(async () => root.unmount())
  })
})

describe('closing a presentation while its slideshow plays', () => {
  it('leaves full screen', async () => {
    setSlidesDocument('deck', readyDocument(newPresentation()))
    startSlideshow('deck')
    const document = readySlides('deck')
    if (document === undefined) throw new Error('No deck')
    const root = mount()
    await act(async () => root.render(<Slideshow documentId="deck" document={document} />))

    releaseSlidesDocument('deck')
    await act(async () => root.unmount())

    expect(setFullScreen.mock.calls).toEqual([[true], [false]])
  })
})

describe('pasting slide objects', () => {
  function clipboard(value: unknown): DataTransfer {
    const json = JSON.stringify(value)
    const data = new DataTransfer()
    data.setData('application/x-zen-suzu-slides+json', json)
    return data
  }

  function lastElement(id: string) {
    const document = readySlides(id)
    return document === undefined ? undefined : currentSlide(document)?.elements.at(-1)
  }

  it('refuses pictures that are not embedded PNG, JPEG or GIF data, and objects missing their parts', () => {
    setSlidesDocument('paste', readyDocument(newPresentation()))
    const before = lastElement('paste')
    paste('paste', clipboard([{ kind: 'image', id: 'a', x: 10, y: 10, width: 10, height: 10, rotation: 0, src: 'https://tracker.example/pixel.png' }]))
    paste('paste', clipboard([{ kind: 'text', x: 10, y: 10 }]))
    expect(lastElement('paste')).toBe(before)
  })

  it('pastes a picture copied from Slides', () => {
    setSlidesDocument('paste2', readyDocument(newPresentation()))
    const src = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP8z8DAwMDAxMDAwMDAAAANHQEDasKb6QAAAABJRU5ErkJggg=='
    paste('paste2', clipboard([{ kind: 'image', id: 'a', x: 10, y: 10, width: 10, height: 10, rotation: 0, src }]))
    expect(lastElement('paste2')).toMatchObject({ kind: 'image', x: 10, src })
  })
})

describe('object commands while typing', () => {
  it('leave the selected object alone when the caret is in its text or in the notes', () => {
    setSlidesDocument('caret', readyDocument(newPresentation()))
    const count = (): number => currentSlide(readySlides('caret') ?? readyDocument(newPresentation()))?.elements.length ?? 0
    const first = readySlides('caret')?.present.slides[0]?.elements[0]
    if (first === undefined) throw new Error('The new presentation has no objects')
    const before = count()
    updateSlides('caret', () => ({ selection: [first.id], editingId: first.id }))
    void runSlideCommand('caret', 'edit.duplicate')
    void runSlideCommand('caret', 'edit.delete')

    const notes = document.createElement('textarea')
    notes.className = 'slides-notes'
    document.body.append(notes)
    cleanups.push(() => notes.remove())
    notes.focus()
    updateSlides('caret', () => ({ editingId: undefined }))
    void runSlideCommand('caret', 'edit.duplicate')
    void runSlideCommand('caret', 'edit.delete')
    expect(count()).toBe(before)

    notes.blur()
    void runSlideCommand('caret', 'edit.duplicate')
    expect(count()).toBe(before + 1)
  })
})

describe('changing a table while a cell is being typed into', () => {
  function Canvas({ id }: { readonly id: string }) {
    const document = useSlidesStore(state => state.documents[id])
    return document?.status === 'ready' ? <SlideCanvas documentId={id} document={document} /> : null
  }

  it('writes the typed text where it was typed before a row is deleted', async () => {
    const start = newPresentation()
    const slideId = start.slides[0]?.id ?? ''
    const table = ['A0', 'B0', 'A1', 'B1'].reduce((current, text, index) => updateCell(current, Math.floor(index / 2), index % 2, { text }), newTable(2, 2, start))
    setSlidesDocument('table', readyDocument(updateSlide(start, slideId, slide => ({ ...slide, elements: [table] }))))
    startEditing('table', table.id, { row: 0, column: 0 })
    const root = mount()
    await act(async () => root.render(<Canvas id="table" />))
    const field = document.querySelector('.slides-cell-editor')
    if (!(field instanceof HTMLTextAreaElement)) throw new Error('No cell editor')

    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')?.set?.call(field, 'draft')
      field.dispatchEvent(new Event('input', { bubbles: true }))
    })
    await act(async () => editTable('table', 'deleteRow'))
    await act(async () => root.unmount())

    const element = readySlides('table')?.present.slides[0]?.elements[0]
    if (element?.kind !== 'table') throw new Error('The table is gone')
    expect(element.rows.map(row => row.cells.map(cell => cell.text))).toEqual([['A1', 'B1']])
  })
})
