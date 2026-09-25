import { useEffect, useSyncExternalStore } from 'react'
import { Editor, type Content, type JSONContent } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Placeholder, Selection } from '@tiptap/extensions'
import { toast } from 'sonner'
import { createImportReport, type ImportFindingInput, type ImportReport } from '@shared/fidelity'
import { useDocumentsStore, type OpenDocument } from '../../app/documents-store'
import { fileService } from '../../services/file/IpcFileService'
import { useFidelityStore } from '../../services/fidelity/fidelity-store'
import { recordPreview, recordRecent } from '../../services/index/document-index'
import { askConfirm } from '../../ui/ConfirmDialog'
import { DocShortcuts } from './doc-commands'
import { FindHighlight } from './find'
import { sameDocument } from './io/compare'
import { localImagePath } from './io/local-images'
import { mimeTypeForExtension, toDataUrl } from './io/images'
import { renderPreview } from './preview'
import { documentExtensions } from './schema'

export type DocEntry =
  | { readonly status: 'loading' }
  | { readonly status: 'ready'; readonly editor: Editor }
  | { readonly status: 'error'; readonly message: string }

interface CachedDoc {
  entry: DocEntry
  /** The document id this editor reports dirtiness to; changes when Save As re-keys the tab. */
  documentId: string
  /** The content as last opened or saved; the tab is dirty while the editor's differs. */
  saved?: ProseMirrorNode
  /** Markdown images shown from data URLs, mapped back to the paths the file had for them. */
  imageSources: Map<string, string>
  /**
   * Set once the person chose Overwrite for a save that would change the file; later saves of this
   * tab then go straight to disk instead of asking every time.
   */
  overwriteAccepted: boolean
}

// Editors outlive the editor component (tabs remount on switch), so they live here, by document id.
const cache = new Map<string, CachedDoc>()
const listeners = new Set<() => void>()
let version = 0

function emit(): void {
  version += 1
  listeners.forEach(listener => listener())
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

function createEditor(content: Content): Editor {
  return new Editor({
    extensions: [
      ...documentExtensions,
      Placeholder.configure({ placeholder: 'Start writing', showOnlyWhenEditable: true }),
      // Keeps the selection visible while the find bar or link field has focus.
      Selection,
      FindHighlight,
      DocShortcuts,
    ],
    content,
    editorProps: { attributes: { class: 'sumi-prose', spellcheck: 'true', 'aria-label': 'Document' } },
  })
}

function setReady(cached: CachedDoc, editor: Editor): void {
  cached.saved = editor.state.doc
  cached.entry = { status: 'ready', editor }
  editor.on('update', () => {
    useDocumentsStore.getState().setDirty(cached.documentId, cached.saved === undefined || !editor.state.doc.eq(cached.saved))
  })
  emit()
}

function directoryOf(path: string): string {
  return path.slice(0, path.lastIndexOf('/') + 1)
}

/**
 * Loads images a Markdown file links by relative path, so they show; the writer puts the original
 * paths back. Images outside the file's folder, or that cannot be read, stay as they are and show as missing.
 */
async function loadLocalImages(node: JSONContent, directory: string, sources: Map<string, string>): Promise<JSONContent> {
  const src: unknown = node.attrs?.['src']
  const path = node.type === 'image' && typeof src === 'string' ? localImagePath(directory, src) : undefined
  if (path !== undefined && typeof src === 'string') {
    const extension = path.split('.').pop()?.toLowerCase() ?? ''
    if (extension === 'png' || extension === 'jpg' || extension === 'jpeg') {
      try {
        const dataUrl = toDataUrl(await fileService.read(path), mimeTypeForExtension(extension))
        sources.set(dataUrl, src)
        return { ...node, attrs: { ...node.attrs, src: dataUrl } }
      } catch {
        return node
      }
    }
  }
  if (node.content === undefined) return node
  return { ...node, content: await Promise.all(node.content.map(child => loadLocalImages(child, directory, sources))) }
}

async function readContent(document: OpenDocument, path: string, cached: CachedDoc): Promise<{ content: Content; findings: readonly ImportFindingInput[] }> {
  const bytes = await fileService.read(path)
  if (document.extension === 'docx') {
    const { readDocx } = await import('./io/docx-read')
    return readDocx(bytes)
  }
  const { readMarkdown } = await import('./io/markdown')
  const imported = readMarkdown(new TextDecoder().decode(bytes))
  return { content: await loadLocalImages(imported.content, directoryOf(path), cached.imageSources), findings: imported.findings }
}

async function load(document: OpenDocument, cached: CachedDoc): Promise<void> {
  if (document.path === undefined) {
    setReady(cached, createEditor(''))
    return
  }
  try {
    const { content, findings } = await readContent(document, document.path, cached)
    // The tab may have closed while the file was read.
    if (cache.get(cached.documentId) !== cached) return
    useFidelityStore.getState().publish(document.id, createImportReport(`${document.name}.${document.extension}`, findings))
    const editor = createEditor(content)
    setReady(cached, editor)
    void recordPreview(document.path, renderPreview(editor.state.doc))
  } catch (error) {
    cached.entry = { status: 'error', message: error instanceof Error ? error.message : 'The file could not be read.' }
    emit()
  }
}

/** The editor for a tab, loading the file on first use. */
export function useDocEntry(document: OpenDocument): DocEntry {
  useSyncExternalStore(subscribe, () => version)
  const cached = cache.get(document.id)
  useEffect(() => {
    if (cache.has(document.id)) return
    const created: CachedDoc = { entry: { status: 'loading' }, documentId: document.id, imageSources: new Map(), overwriteAccepted: false }
    cache.set(document.id, created)
    void load(document, created)
  }, [document])
  return cached?.entry ?? { status: 'loading' }
}

export function docEditor(documentId: string): Editor | undefined {
  const entry = cache.get(documentId)?.entry
  return entry?.status === 'ready' ? entry.editor : undefined
}

export function releaseDoc(documentId: string): void {
  const cached = cache.get(documentId)
  if (cached?.entry.status === 'ready') cached.entry.editor.destroy()
  cache.delete(documentId)
  useFidelityStore.getState().forget(documentId)
  emit()
}

function extensionOf(path: string): string {
  return path.split('.').pop()?.toLowerCase() ?? ''
}

/** Save As offers the document's own format first; new documents default to .docx. */
function saveFormats(extension: string): readonly string[] {
  return extension === 'md' || extension === 'markdown' ? [extension, 'docx'] : ['docx', 'md']
}

function isMarkdown(format: string): boolean {
  return format === 'md' || format === 'markdown'
}

async function encode(editor: Editor, format: string, imageSources: ReadonlyMap<string, string>): Promise<Uint8Array> {
  if (isMarkdown(format)) {
    const { writeMarkdown } = await import('./io/markdown')
    return new TextEncoder().encode(writeMarkdown(editor.getJSON(), imageSources))
  }
  const { writeDocx } = await import('./io/docx-export')
  const { bytes, skippedImages } = await writeDocx(editor.getJSON())
  if (skippedImages > 0) {
    toast.warning(`${skippedImages} ${skippedImages === 1 ? 'image was' : 'images were'} not saved`, {
      description: 'Word files can hold PNG, JPEG, GIF and BMP images that are part of the document, not linked from elsewhere.',
    })
  }
  return bytes
}

/** The document with each image a Markdown file linked by path back at that path, as the file holds it. */
function withImagePaths(node: JSONContent, sources: ReadonlyMap<string, string>): JSONContent {
  const src: unknown = node.attrs?.['src']
  const original = node.type === 'image' && typeof src === 'string' ? sources.get(src) : undefined
  return {
    ...node,
    ...(original === undefined ? {} : { attrs: { ...node.attrs, src: original } }),
    ...(node.content === undefined ? {} : { content: node.content.map(child => withImagePaths(child, sources)) }),
  }
}

/**
 * Reads `bytes` back with Sumi's own reader and compares the result with the editor's document.
 * Anything a format cannot hold, or that the writer and reader disagree on, shows up here, so a
 * save never replaces a file with less than the editor shows without asking.
 */
async function changesOnReopen(editor: Editor, format: string, bytes: Uint8Array, imageSources: ReadonlyMap<string, string>): Promise<boolean> {
  try {
    if (isMarkdown(format)) {
      const { readMarkdown } = await import('./io/markdown')
      return !sameDocument(editor.schema, withImagePaths(editor.getJSON(), imageSources), readMarkdown(new TextDecoder().decode(bytes)).content)
    }
    const { readDocx } = await import('./io/docx-read')
    return !sameDocument(editor.schema, editor.getJSON(), readDocx(bytes).content)
  } catch {
    return true
  }
}

/** Why saving over the file would lose something, or undefined when it keeps everything. */
async function overwriteRisk(editor: Editor, format: string, bytes: Uint8Array, cached: CachedDoc, report: ImportReport | undefined): Promise<string | undefined> {
  if (report !== undefined && report.severity !== 'lossless') {
    return `Saving writes only what Sumi shows, so what the import report lists would change or be lost in ${report.sourceName}.`
  }
  if (await changesOnReopen(editor, format, bytes, cached.imageSources)) {
    return isMarkdown(format) ? 'Saving as Markdown would change parts of this document.' : 'Saving as a Word document would change parts of this document.'
  }
  return undefined
}

/**
 * Saves a document tab. Saving over an existing file first checks that nothing would be lost: the
 * import report is clean, and the file Sumi writes reads back as the document shown. Otherwise it
 * asks: Overwrite, or Save a copy, which keeps the original intact.
 */
export async function saveDoc(documentId: string, saveAs: boolean): Promise<boolean> {
  const document = useDocumentsStore.getState().documents.find(candidate => candidate.id === documentId)
  const cached = cache.get(documentId)
  const editor = docEditor(documentId)
  if (document === undefined || cached === undefined || editor === undefined) return false

  const imported = useFidelityStore.getState().reports[documentId]
  let path = saveAs ? undefined : document.path
  let bytes: Uint8Array | undefined
  if (path !== undefined && !cached.overwriteAccepted) {
    bytes = await encode(editor, extensionOf(path), cached.imageSources)
    const risk = await overwriteRisk(editor, extensionOf(path), bytes, cached, imported)
    if (risk !== undefined) {
      const { value } = await askConfirm({
        title: 'Save over the original?',
        message: `${risk} Saving a copy keeps the original file intact.`,
        actions: [{ label: 'Overwrite', value: 'overwrite' }, { label: 'Save a copy', value: 'copy', primary: true }],
      })
      if (value === 'cancel') return false
      if (value === 'copy') {
        path = undefined
        bytes = undefined
      } else cached.overwriteAccepted = true
    }
  }

  if (path === undefined) {
    const formats = saveFormats(document.extension)
    const suffix = !saveAs && document.path !== undefined ? ' (edited)' : ''
    path = await fileService.chooseSavePath({ defaultName: `${document.name}${suffix}.${formats[0] ?? 'docx'}`, extensions: formats })
    if (path === undefined) return false
  }

  const format = extensionOf(path)
  const snapshot = editor.state.doc
  await fileService.write(path, bytes ?? await encode(editor, format, cached.imageSources))
  const file = await fileService.describe(path)
  cached.saved = snapshot

  const newId = useDocumentsStore.getState().setSaved(documentId, file)
  if (newId !== documentId) {
    cache.delete(documentId)
    cached.documentId = newId
    cached.overwriteAccepted = false
    cache.set(newId, cached)
    emit()
  }
  // Edits made while the file was being written keep the tab dirty.
  useDocumentsStore.getState().setDirty(newId, !editor.state.doc.eq(snapshot))
  // The file on disk is now one Sumi wrote, so it holds nothing the editor does not show.
  useFidelityStore.getState().forget(documentId)
  if (imported !== undefined) useFidelityStore.getState().publish(newId, createImportReport(file.name, []))
  await recordRecent(file)
  await recordPreview(file.path, renderPreview(snapshot))
  toast.success(`Saved ${file.name}`)
  if (isMarkdown(format) && !isMarkdown(document.extension)) {
    toast.info('Markdown keeps text and structure only', { description: 'Fonts, colours, underline, alignment, spacing, page setup and page breaks are saved in .docx.' })
  }
  return true
}
