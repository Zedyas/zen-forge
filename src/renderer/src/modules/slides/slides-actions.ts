import { toast } from 'sonner'
import type { CommandId } from '@shared/commands'
import { createImportReport } from '@shared/fidelity'
import { useDocumentsStore } from '../../app/documents-store'
import { fileService } from '../../services/file/IpcFileService'
import { useFidelityStore } from '../../services/fidelity/fidelity-store'
import { recordRecent } from '../../services/index/document-index'
import { askConfirm } from '../pdf/ConfirmDialog'
import {
  allRuns,
  arrange,
  createSlide,
  duplicateElement,
  duplicateSlide,
  formatRuns,
  holdsText,
  newImage,
  newPresentation,
  translateElement,
  updateElement,
  updateSlide,
  type HorizontalAlign,
  type Presentation,
  type RunStyle,
  type Slide,
  type SlideElement,
  type SlideLayout,
} from './model'
import {
  commitSlides,
  markSlidesSaved,
  readyDocument,
  readySlides,
  redoSlides,
  rekeySlides,
  releaseSlides,
  selectedElement,
  setSlidesDocument,
  undoSlides,
  updateSlides,
} from './slides-store'
import { activeEditor } from './text-editing'

function errorMessage(error: unknown, fallback: string): string {
  return error instanceof Error && error.message !== '' ? error.message : fallback
}

/* ─── Load, save, release ─── */

/** The latest load of each document. A load that is no longer the latest, because its tab closed or reopened, drops its result. */
const loads = new Map<string, symbol>()

/** Loads a .pptx into the store and publishes its import report. */
export async function loadSlidesDocument(id: string, path: string, fileName: string): Promise<void> {
  const load = Symbol(id)
  loads.set(id, load)
  setSlidesDocument(id, { status: 'loading' })
  try {
    const bytes = await fileService.read(path)
    const { readPptx } = await import('./pptx-import')
    const { presentation, findings } = await readPptx(bytes)
    if (loads.get(id) !== load) return
    setSlidesDocument(id, readyDocument(presentation))
    useFidelityStore.getState().publish(id, createImportReport(fileName, findings))
  } catch (error) {
    if (loads.get(id) === load) setSlidesDocument(id, { status: 'error', message: errorMessage(error, 'The file could not be read.') })
  }
}

export function startNewPresentation(id: string): void {
  setSlidesDocument(id, readyDocument(newPresentation()))
}

export function releaseSlidesDocument(id: string): void {
  loads.delete(id)
  releaseSlides(id)
  useFidelityStore.getState().forget(id)
}

/** Closes the inline text editor and the notes field, writing what was typed, so a save or change includes it. */
export function finishTyping(id: string): void {
  activeEditor(id)?.commit()
  if (window.document.activeElement instanceof HTMLElement) window.document.activeElement.blur()
}

/**
 * Saves in place, or asks for a path. A presentation opened from a file that Zendo could not keep
 * exactly asks first whether to overwrite it or save a copy.
 */
export async function saveSlidesDocument(id: string, saveAs: boolean): Promise<boolean> {
  finishTyping(id)
  const document = readySlides(id)
  const open = useDocumentsStore.getState().documents.find(candidate => candidate.id === id)
  if (document === undefined || open === undefined) return false

  const report = useFidelityStore.getState().reports[id]
  let chooseNewPath = saveAs || open.path === undefined
  let suffix = ''
  if (!chooseNewPath && report !== undefined && report.severity !== 'lossless') {
    const { value } = await askConfirm({
      title: 'Save over the original?',
      message: report.severity === 'dropped'
        ? 'This presentation contains content that saving removes (listed in the import report). Saving a copy keeps the original file intact.'
        : 'Some of this presentation is saved differently from the original (listed in the import report). Saving a copy keeps the original file as it was.',
      actions: [{ label: 'Overwrite', value: 'overwrite' }, { label: 'Save a copy', value: 'copy', primary: true }],
    })
    if (value === 'cancel') return false
    chooseNewPath = value === 'copy'
    if (chooseNewPath) suffix = ' (edited)'
  }

  const path = chooseNewPath
    ? await fileService.chooseSavePath({ defaultName: `${open.name}${suffix}.pptx`, extensions: ['pptx'] })
    : open.path
  if (path === undefined) return false

  const presentation = document.present
  const { writePptx } = await import('./pptx-export')
  await fileService.write(path, await writePptx(presentation))
  const file = await fileService.describe(path)
  const nextId = useDocumentsStore.getState().setSaved(id, file)
  rekeySlides(id, nextId)
  markSlidesSaved(nextId, presentation)
  // The file on disk is now one Zendo wrote, so it holds nothing unrepresented and later saves need no question.
  if (report !== undefined) {
    useFidelityStore.getState().forget(id)
    useFidelityStore.getState().publish(nextId, createImportReport(file.name, []))
  }
  await recordRecent(file)
  toast.success('Saved', { description: file.name })
  return true
}

/* ─── Slides ─── */

export function activeSlidesId(): string | undefined {
  const { activeId } = useDocumentsStore.getState()
  return readySlides(activeId) === undefined ? undefined : activeId
}

export function showSlide(id: string, slideId: string): void {
  finishTyping(id)
  updateSlides(id, () => ({ slideId, selectedId: undefined, editingId: undefined }))
}

function currentIndex(id: string): number {
  const document = readySlides(id)
  return document === undefined ? -1 : document.present.slides.findIndex(slide => slide.id === document.slideId)
}

export function addSlide(id: string, layout: SlideLayout): void {
  finishTyping(id)
  const document = readySlides(id)
  if (document === undefined) return
  const slide = createSlide(layout, document.present)
  const at = currentIndex(id) + 1
  commitSlides(id, presentation => ({ ...presentation, slides: [...presentation.slides.slice(0, at), slide, ...presentation.slides.slice(at)] }))
  updateSlides(id, () => ({ slideId: slide.id, selectedId: undefined }))
}

export function duplicateCurrentSlide(id: string, slideId?: string): void {
  finishTyping(id)
  const document = readySlides(id)
  const source = document?.present.slides.find(slide => slide.id === (slideId ?? document.slideId))
  if (source === undefined) return
  const copy = duplicateSlide(source)
  commitSlides(id, presentation => {
    const at = presentation.slides.findIndex(slide => slide.id === source.id) + 1
    return { ...presentation, slides: [...presentation.slides.slice(0, at), copy, ...presentation.slides.slice(at)] }
  })
  updateSlides(id, () => ({ slideId: copy.id, selectedId: undefined }))
}

export function deleteSlide(id: string, slideId?: string): void {
  finishTyping(id)
  const document = readySlides(id)
  if (document === undefined) return
  const target = slideId ?? document.slideId
  if (document.present.slides.length === 1) {
    toast.error('A presentation needs at least one slide.')
    return
  }
  const index = document.present.slides.findIndex(slide => slide.id === target)
  const remaining = document.present.slides.filter(slide => slide.id !== target)
  const next = document.slideId === target ? remaining[Math.min(index, remaining.length - 1)]?.id : document.slideId
  commitSlides(id, presentation => ({ ...presentation, slides: remaining }))
  updateSlides(id, () => ({ slideId: next ?? '', selectedId: undefined }))
}

export function moveSlide(id: string, fromId: string, toId: string): void {
  commitSlides(id, presentation => {
    const from = presentation.slides.findIndex(slide => slide.id === fromId)
    const to = presentation.slides.findIndex(slide => slide.id === toId)
    if (from < 0 || to < 0 || from === to) return presentation
    const slides = [...presentation.slides]
    const [moved] = slides.splice(from, 1)
    if (moved !== undefined) slides.splice(to, 0, moved)
    return { ...presentation, slides }
  })
}

export function setSlideNotes(id: string, slideId: string, notes: string): void {
  commitSlides(id, presentation => {
    const slide = presentation.slides.find(candidate => candidate.id === slideId)
    return slide === undefined || slide.notes === notes ? presentation : updateSlide(presentation, slideId, current => ({ ...current, notes }))
  })
}

export function setSlideBackground(id: string, background: string): void {
  const document = readySlides(id)
  if (document === undefined) return
  commitSlides(id, presentation => updateSlide(presentation, document.slideId, slide => ({ ...slide, background })))
}

/* ─── Elements ─── */

/** Adds an element to the current slide as one undoable step and selects it; `edit` opens its text for typing. */
export function addElement(id: string, element: SlideElement, edit = false): void {
  finishTyping(id)
  const document = readySlides(id)
  if (document === undefined) return
  commitSlides(id, presentation => updateSlide(presentation, document.slideId, slide => ({ ...slide, elements: [...slide.elements, element] })))
  updateSlides(id, () => ({ selectedId: element.id, editingId: edit ? element.id : undefined }))
}

/** Changes an element on whichever slide holds it, so a text editor closing after a slide change still lands. */
function onSlideHolding(presentation: Presentation, elementId: string, change: (slide: Slide) => Slide): Presentation {
  const slide = presentation.slides.find(candidate => candidate.elements.some(element => element.id === elementId))
  return slide === undefined ? presentation : updateSlide(presentation, slide.id, change)
}

export function changeElement(id: string, elementId: string, change: (element: SlideElement) => SlideElement): void {
  commitSlides(id, presentation => onSlideHolding(presentation, elementId, slide => updateElement(slide, elementId, change)))
}

export function removeElement(id: string, elementId: string): void {
  commitSlides(id, presentation => onSlideHolding(presentation, elementId, slide => ({
    ...slide,
    elements: slide.elements.filter(element => element.id !== elementId),
  })))
  updateSlides(id, current => current.selectedId === elementId ? { selectedId: undefined, editingId: undefined } : {})
}

export function select(id: string, elementId: string | undefined): void {
  const document = readySlides(id)
  if (document === undefined || (document.selectedId === elementId && document.editingId === undefined)) return
  if (document.editingId !== undefined && document.editingId !== elementId) activeEditor(id)?.commit()
  updateSlides(id, () => ({ selectedId: elementId, editingId: undefined }))
}

export function startEditing(id: string, elementId: string): void {
  updateSlides(id, () => ({ selectedId: elementId, editingId: elementId }))
}

export function duplicateSelection(id: string): void {
  const document = readySlides(id)
  const element = document === undefined ? undefined : selectedElement(document)
  if (element !== undefined) addElement(id, duplicateElement(element, 12))
}

export function arrangeSelection(id: string, to: 'front' | 'back'): void {
  const document = readySlides(id)
  if (document?.selectedId === undefined) return
  const elementId = document.selectedId
  commitSlides(id, presentation => updateSlide(presentation, document.slideId, slide => arrange(slide, elementId, to)))
}

export function nudgeSelection(id: string, dx: number, dy: number): void {
  const document = readySlides(id)
  if (document?.selectedId !== undefined) changeElement(id, document.selectedId, element => translateElement(element, dx, dy))
}

/* ─── Text formatting: the selected text while typing, otherwise the whole selected box ─── */

function selectedText(id: string) {
  const document = readySlides(id)
  const element = document === undefined ? undefined : selectedElement(document)
  return element !== undefined && holdsText(element) ? element : undefined
}

export function applyRunStyle(id: string, change: Partial<RunStyle>): void {
  const editor = activeEditor(id)
  if (editor !== undefined) {
    if (change.color !== undefined) editor.format({ kind: 'color', value: change.color })
    if (change.size !== undefined) editor.format({ kind: 'size', value: change.size })
    return
  }
  const element = selectedText(id)
  if (element !== undefined) changeElement(id, element.id, current => holdsText(current) ? { ...current, paragraphs: formatRuns(current.paragraphs, change) } : current)
}

export function toggleRunStyle(id: string, style: 'bold' | 'italic' | 'underline'): void {
  const editor = activeEditor(id)
  if (editor !== undefined) {
    editor.format({ kind: 'toggle', style })
    return
  }
  const element = selectedText(id)
  if (element === undefined) return
  const runs = allRuns(element.paragraphs)
  applyRunStyle(id, { [style]: !(runs.length > 0 && runs.every(run => run[style])) })
}

export function setAlign(id: string, align: HorizontalAlign): void {
  const editor = activeEditor(id)
  if (editor !== undefined) {
    editor.format({ kind: 'align', value: align })
    return
  }
  const element = selectedText(id)
  if (element !== undefined) {
    changeElement(id, element.id, current => holdsText(current)
      ? { ...current, paragraphs: current.paragraphs.map(paragraph => ({ ...paragraph, align })) }
      : current)
  }
}

export function setFill(id: string, fill: string | undefined): void {
  const document = readySlides(id)
  const element = document === undefined ? undefined : selectedElement(document)
  if (element === undefined || element.kind === 'image') return
  changeElement(id, element.id, current => current.kind === 'image' ? current : { ...current, fill })
}

/* ─── Pictures ─── */

function dataUrl(bytes: Uint8Array, type: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => typeof reader.result === 'string' ? resolve(reader.result) : reject(new Error('The picture could not be read.'))
    reader.onerror = () => reject(reader.error ?? new Error('The picture could not be read.'))
    reader.readAsDataURL(new Blob([new Uint8Array(bytes)], { type }))
  })
}

/** Asks for a PNG or JPEG and places it in the middle of the current slide. */
export async function insertPicture(id: string): Promise<void> {
  const [file] = await fileService.chooseFiles({ extensions: ['png', 'jpg', 'jpeg'] })
  const document = readySlides(id)
  if (file === undefined || document === undefined) return
  const src = await dataUrl(await fileService.read(file.path), file.extension === 'png' ? 'image/png' : 'image/jpeg')
  const image = new Image()
  image.src = src
  await image.decode()
  // Pixels at 96 per inch, as PowerPoint places pictures without a resolution.
  addElement(id, newImage(src, { width: image.naturalWidth * 0.75, height: image.naturalHeight * 0.75 }, document.present))
}

/* ─── Commands ─── */

export function isTextEntry(target: EventTarget | null): boolean {
  return target instanceof HTMLElement && target.matches('input, textarea, select, [contenteditable="true"]')
}

/** Editor commands from the menu, keyboard and toolbar, applied to the active presentation. */
export async function runSlidesCommand(command: CommandId): Promise<void> {
  const id = activeSlidesId()
  if (id === undefined) return
  try {
    switch (command) {
      case 'edit.undo':
        // A focused field keeps its own text undo; the menu accelerator would otherwise swallow it.
        if (isTextEntry(window.document.activeElement)) window.document.execCommand('undo')
        else undoSlides(id)
        return
      case 'edit.redo':
        if (isTextEntry(window.document.activeElement)) window.document.execCommand('redo')
        else redoSlides(id)
        return
      case 'format.bold':
        return toggleRunStyle(id, 'bold')
      case 'format.italic':
        return toggleRunStyle(id, 'italic')
      case 'format.underline':
        return toggleRunStyle(id, 'underline')
      case 'slide.new':
        return addSlide(id, 'titleContent')
      case 'slide.duplicate':
        return duplicateCurrentSlide(id)
      case 'slide.delete':
        return deleteSlide(id)
      default:
        return
    }
  } catch (error) {
    toast.error('Could not complete that', { description: errorMessage(error, 'The presentation could not be changed.') })
  }
}

