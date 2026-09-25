import { toast } from 'sonner'
import type { CommandId } from '@shared/commands'
import { createImportReport } from '@shared/fidelity'
import { useDocumentsStore } from '../../app/documents-store'
import { fileService } from '../../services/file/IpcFileService'
import { useFidelityStore } from '../../services/fidelity/fidelity-store'
import { recordRecent } from '../../services/index/document-index'
import { platformClient } from '../../services/platform/client'
import { askConfirm } from '../pdf/ConfirmDialog'
import {
  alignTo,
  applyTheme,
  arrange,
  boundsOf,
  createSlide,
  duplicateSlide,
  newId,
  newImage,
  newPresentation,
  mapElements,
  newTextBox,
  paragraphOf,
  translateElement,
  updateElement,
  updateSlide,
  type AlignEdge,
  type Presentation,
  type Slide,
  type SlideElement,
  type SlideLayout,
} from './model'
import {
  commitSlides,
  currentSlide,
  markSlidesSaved,
  readyDocument,
  readySlides,
  redoSlides,
  rekeySlides,
  releaseSlides,
  selectedElements,
  setSlidesDocument,
  undoSlides,
  updateSlides,
  type TableCellAddress,
} from './slides-store'
import { activeEditor } from './text-editing'
import { findTheme } from './themes'

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
  updateSlides(id, () => ({ slideId, selection: [], editingId: undefined, cell: undefined }))
}

function currentIndex(id: string): number {
  const document = readySlides(id)
  return document === undefined ? -1 : document.present.slides.findIndex(slide => slide.id === document.slideId)
}

function insertSlideAfter(presentation: Presentation, afterId: string, slide: Slide): Presentation {
  const at = presentation.slides.findIndex(candidate => candidate.id === afterId) + 1
  return { ...presentation, slides: [...presentation.slides.slice(0, at), slide, ...presentation.slides.slice(at)] }
}

export function addSlide(id: string, layout: SlideLayout): void {
  finishTyping(id)
  const document = readySlides(id)
  if (document === undefined) return
  const slide = createSlide(layout, document.present)
  commitSlides(id, presentation => insertSlideAfter(presentation, document.slideId, slide))
  updateSlides(id, () => ({ slideId: slide.id, selection: [] }))
}

export function duplicateCurrentSlide(id: string, slideId?: string): void {
  finishTyping(id)
  const document = readySlides(id)
  const source = document?.present.slides.find(slide => slide.id === (slideId ?? document.slideId))
  if (source === undefined) return
  const copy = duplicateSlide(source)
  commitSlides(id, presentation => insertSlideAfter(presentation, source.id, copy))
  updateSlides(id, () => ({ slideId: copy.id, selection: [] }))
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
  updateSlides(id, () => ({ slideId: next ?? '', selection: [] }))
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

/** Changes the current slide's background, or every slide's. */
export function setBackground(id: string, change: Pick<Slide, 'background' | 'backgroundImage'>, everySlide = false): void {
  const document = readySlides(id)
  if (document === undefined) return
  const same = (slide: Slide): boolean => slide.background === change.background && slide.backgroundImage === change.backgroundImage
  commitSlides(id, presentation => presentation.slides.every(slide => !(everySlide || slide.id === document.slideId) || same(slide))
    ? presentation
    : { ...presentation, slides: presentation.slides.map(slide => everySlide || slide.id === document.slideId ? { ...slide, ...change } : slide) })
}

export function setTheme(id: string, themeId: string): void {
  commitSlides(id, presentation => presentation.theme === themeId ? presentation : applyTheme(presentation, findTheme(themeId)))
}

/* ─── Selection ─── */

export function select(id: string, elementIds: readonly string[]): void {
  const document = readySlides(id)
  if (document === undefined) return
  if (document.editingId !== undefined && !(elementIds.length === 1 && elementIds[0] === document.editingId)) activeEditor(id)?.commit()
  const same = elementIds.length === document.selection.length && elementIds.every((elementId, index) => document.selection[index] === elementId)
  if (same && document.editingId === undefined) return
  updateSlides(id, current => ({ selection: elementIds, editingId: undefined, cell: same ? current.cell : undefined }))
}

/** Shift-click: adds an element to the selection, or takes it out. */
export function toggleSelected(id: string, elementId: string): void {
  const document = readySlides(id)
  if (document === undefined) return
  select(id, document.selection.includes(elementId) ? document.selection.filter(candidate => candidate !== elementId) : [...document.selection, elementId])
}

/** Opens an element's text for typing; for a table, the text of one cell. */
export function startEditing(id: string, elementId: string, cell?: TableCellAddress): void {
  updateSlides(id, () => ({ selection: [elementId], editingId: elementId, cell }))
}

export function setCurrentCell(id: string, cell: TableCellAddress | undefined): void {
  updateSlides(id, () => ({ cell }))
}

/* ─── Elements ─── */

/** Adds elements to the current slide as one undoable step and selects them; `edit` opens the first for typing. */
export function addElements(id: string, elements: readonly SlideElement[], edit = false): void {
  finishTyping(id)
  const document = readySlides(id)
  const first = elements[0]
  if (document === undefined || first === undefined) return
  commitSlides(id, presentation => updateSlide(presentation, document.slideId, slide => ({ ...slide, elements: [...slide.elements, ...elements] })))
  updateSlides(id, () => ({ selection: elements.map(element => element.id), editingId: edit ? first.id : undefined, cell: undefined }))
}

export function addElement(id: string, element: SlideElement, edit = false): void {
  addElements(id, [element], edit)
}

/** Changes an element on whichever slide holds it, so a text editor closing after a slide change still lands. */
function onSlideHolding(presentation: Presentation, elementId: string, change: (slide: Slide) => Slide): Presentation {
  const slide = presentation.slides.find(candidate => candidate.elements.some(element => element.id === elementId))
  return slide === undefined ? presentation : updateSlide(presentation, slide.id, change)
}

export function changeElement(id: string, elementId: string, change: (element: SlideElement) => SlideElement): void {
  commitSlides(id, presentation => onSlideHolding(presentation, elementId, slide => updateElement(slide, elementId, change)))
}

/** Changes several elements of the current slide as one undoable step. */
export function changeElements(id: string, elementIds: readonly string[], change: (element: SlideElement) => SlideElement): void {
  const document = readySlides(id)
  if (document === undefined || elementIds.length === 0) return
  const ids = new Set(elementIds)
  commitSlides(id, presentation => updateSlide(presentation, document.slideId, slide => mapElements(slide, element => ids.has(element.id) ? change(element) : element)))
}

/** Replaces elements of the current slide with changed copies, as a drag ends. */
export function replaceElements(id: string, changed: readonly SlideElement[]): void {
  const byId = new Map(changed.map(element => [element.id, element]))
  changeElements(id, [...byId.keys()], element => byId.get(element.id) ?? element)
}

export function removeElements(id: string, elementIds: readonly string[]): void {
  const document = readySlides(id)
  if (document === undefined) return
  const ids = new Set(elementIds)
  commitSlides(id, presentation => ({
    ...presentation,
    slides: presentation.slides.map(slide => slide.elements.some(element => ids.has(element.id))
      ? { ...slide, elements: slide.elements.filter(element => !ids.has(element.id)) }
      : slide),
  }))
}

export function removeSelection(id: string): void {
  const document = readySlides(id)
  if (document !== undefined) removeElements(id, document.selection)
}

/** Copies of elements with new ids, shifted when pasted where the originals sit. */
function copiesFor(slide: Slide | undefined, elements: readonly SlideElement[]): SlideElement[] {
  const taken = new Set(slide?.elements.map(element => `${Math.round(element.x)},${Math.round(element.y)}`))
  const offset = elements.some(element => taken.has(`${Math.round(element.x)},${Math.round(element.y)}`)) ? 12 : 0
  return elements.map(element => ({ ...translateElement(element, offset, offset), id: newId() }))
}

export function duplicateSelection(id: string): void {
  const document = readySlides(id)
  if (document === undefined) return
  addElements(id, copiesFor(currentSlide(document), selectedElements(document)))
}

export function arrangeSelection(id: string, to: 'front' | 'back'): void {
  const document = readySlides(id)
  if (document === undefined || document.selection.length === 0) return
  const ids = new Set(document.selection)
  commitSlides(id, presentation => updateSlide(presentation, document.slideId, slide => arrange(slide, ids, to)))
}

export function nudgeSelection(id: string, dx: number, dy: number): void {
  const document = readySlides(id)
  if (document !== undefined) changeElements(id, document.selection, element => translateElement(element, dx, dy))
}

/** Lines selected elements up with each other, or a single element with the slide. */
export function alignSelection(id: string, edge: AlignEdge): void {
  const document = readySlides(id)
  if (document === undefined) return
  const elements = selectedElements(document)
  const target = elements.length > 1 ? boundsOf(elements) : { x: 0, y: 0, width: document.present.width, height: document.present.height }
  if (target !== undefined) changeElements(id, document.selection, element => alignTo(element, edge, target))
}

/* ─── Clipboard: elements travel as JSON in a private type, so they paste into any Zendo window ─── */

const clipboardType = 'application/x-zendo-slides+json'
const elementKinds: ReadonlySet<string> = new Set(['text', 'shape', 'image', 'table'])

/** Pasted JSON comes from Zendo's own copy, but it passed through the system clipboard, so its shape is checked. */
function isSlideElement(value: unknown): value is SlideElement {
  return typeof value === 'object' && value !== null && 'kind' in value && typeof value.kind === 'string' && elementKinds.has(value.kind)
    && 'x' in value && typeof value.x === 'number' && 'y' in value && typeof value.y === 'number'
}

function readClipboardElements(data: DataTransfer): SlideElement[] | undefined {
  const json = data.getData(clipboardType)
  if (json === '') return undefined
  try {
    const parsed: unknown = JSON.parse(json)
    return Array.isArray(parsed) && parsed.every(isSlideElement) ? parsed : undefined
  } catch {
    return undefined
  }
}

/** Puts the selected elements on the clipboard; `cut` also removes them. Resolves false when nothing is selected. */
export function copySelection(id: string, data: DataTransfer, cut: boolean): boolean {
  const document = readySlides(id)
  const elements = document === undefined ? [] : selectedElements(document)
  if (elements.length === 0) return false
  data.setData(clipboardType, JSON.stringify(elements))
  data.setData('text/plain', elements.flatMap(element => 'paragraphs' in element ? element.paragraphs.map(paragraph => paragraph.runs.map(run => run.text).join('')) : []).join('\n'))
  if (cut) removeElements(id, elements.map(element => element.id))
  return true
}

/** Pastes copied elements, a picture, or plain text as a new text box. Resolves false when the clipboard holds none of those. */
export function paste(id: string, data: DataTransfer): boolean {
  const document = readySlides(id)
  if (document === undefined) return false
  const elements = readClipboardElements(data)
  if (elements !== undefined) {
    addElements(id, copiesFor(currentSlide(document), elements))
    return true
  }
  const picture = Array.from(data.files).find(file => file.type === 'image/png' || file.type === 'image/jpeg')
  if (picture !== undefined) {
    void picture.arrayBuffer().then(buffer => placePicture(id, new Uint8Array(buffer), picture.type)).catch(() => toast.error('Could not paste the picture'))
    return true
  }
  const text = data.getData('text/plain')
  if (text.trim() === '') return false
  const box = newTextBox(document.present)
  const style = box.paragraphs[0]?.runs[0]
  addElement(id, { ...box, paragraphs: text.split(/\r?\n/).map(line => paragraphOf(line, style)) })
  return true
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

/** Asks for a PNG or JPEG; resolves its data URL and size in points, or undefined when cancelled. */
async function choosePicture(): Promise<{ readonly src: string; readonly width: number; readonly height: number } | undefined> {
  const [file] = await fileService.chooseFiles({ extensions: ['png', 'jpg', 'jpeg'] })
  if (file === undefined) return undefined
  return pictureOf(await fileService.read(file.path), file.extension === 'png' ? 'image/png' : 'image/jpeg')
}

async function pictureOf(bytes: Uint8Array, type: string) {
  const src = await dataUrl(bytes, type)
  const image = new Image()
  image.src = src
  await image.decode()
  // Pixels at 96 per inch, as PowerPoint places pictures without a resolution.
  return { src, width: image.naturalWidth * 0.75, height: image.naturalHeight * 0.75 }
}

async function placePicture(id: string, bytes: Uint8Array, type: string): Promise<void> {
  const picture = await pictureOf(bytes, type)
  const document = readySlides(id)
  if (document !== undefined) addElement(id, newImage(picture.src, picture, document.present))
}

/** Asks for a PNG or JPEG and places it in the middle of the current slide. */
export async function insertPicture(id: string): Promise<void> {
  const picture = await choosePicture()
  const document = readySlides(id)
  if (picture !== undefined && document !== undefined) addElement(id, newImage(picture.src, picture, document.present))
}

export async function chooseBackgroundPicture(id: string): Promise<void> {
  const picture = await choosePicture()
  const slide = readySlides(id)
  const current = slide === undefined ? undefined : currentSlide(slide)
  if (picture !== undefined && current !== undefined) setBackground(id, { background: current.background, backgroundImage: picture.src })
}

/* ─── Slideshow ─── */

export function startSlideshow(id: string): void {
  finishTyping(id)
  const index = Math.max(0, currentIndex(id))
  updateSlides(id, () => ({ playing: index, selection: [], editingId: undefined }))
  void platformClient.setFullScreen(true)
}

/** Ends the slideshow on the slide it was showing, as Keynote does. */
export function stopSlideshow(id: string): void {
  updateSlides(id, document => ({
    playing: undefined,
    slideId: document.playing === undefined ? document.slideId : document.present.slides[document.playing]?.id ?? document.slideId,
  }))
  void platformClient.setFullScreen(false)
}

/** Steps the slideshow; stepping past the last slide ends it. */
export function stepSlideshow(id: string, to: number): void {
  const document = readySlides(id)
  if (document?.playing === undefined) return
  if (to >= document.present.slides.length) {
    stopSlideshow(id)
    return
  }
  updateSlides(id, () => ({ playing: Math.max(0, to) }))
}

/* ─── Commands ─── */

export function isTextEntry(target: EventTarget | null): boolean {
  return target instanceof HTMLElement && target.matches('input, textarea, select, [contenteditable="true"]')
}

/** Editor commands from the menu, keyboard and toolbar that are not about text formatting (see format-actions). */
export function runSlideCommand(id: string, command: CommandId): void | Promise<void> {
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
    case 'slide.new':
      return addSlide(id, 'titleContent')
    case 'slide.duplicate':
      return duplicateCurrentSlide(id)
    case 'slide.delete':
      return deleteSlide(id)
    case 'slide.play':
      return startSlideshow(id)
    default:
      return
  }
}
