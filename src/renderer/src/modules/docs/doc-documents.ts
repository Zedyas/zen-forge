import { useEffect, useSyncExternalStore } from 'react'
import { Editor, type Content, type JSONContent } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Placeholder, Selection } from '@tiptap/extensions'
import { toast } from 'sonner'
import { createImportReport, type ImportFindingInput } from '@shared/fidelity'
import { useDocumentsStore, type OpenDocument } from '../../app/documents-store'
import { fileService } from '../../services/file/IpcFileService'
import { useFidelityStore } from '../../services/fidelity/fidelity-store'
import { recordPreview, recordRecent } from '../../services/index/document-index'
import { askConfirm } from '../pdf/ConfirmDialog'
import { DocShortcuts } from './doc-commands'
import { FindHighlight } from './find'
import { mimeTypeForExtension, toDataUrl } from './io/images'
import { localPageFormat } from './page'
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
}

/** One page format for the whole app run: the one the editor shows is the one .docx files get. */
export const page = localPageFormat()

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

/** Resolves `images/a.png` or `../a.png` against the folder of the Markdown file. */
function resolveRelative(directory: string, src: string): string {
  const segments = (src.startsWith('/') ? src : `${directory}${src}`).split('/')
  const resolved: string[] = []
  for (const segment of segments) {
    if (segment === '..') resolved.pop()
    else if (segment !== '.') resolved.push(segment)
  }
  return decodeURI(resolved.join('/'))
}

/**
 * Loads images a Markdown file links by relative path, so they show; the writer puts the original
 * paths back. Images that cannot be read stay as they are and show as missing.
 */
async function loadLocalImages(node: JSONContent, directory: string, sources: Map<string, string>): Promise<JSONContent> {
  const src: unknown = node.attrs?.['src']
  if (node.type === 'image' && typeof src === 'string' && !/^[a-z][a-z0-9+.-]*:/i.test(src)) {
    const path = resolveRelative(directory, src)
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
    const { readDocx } = await import('./io/docx-import')
    const imported = await readDocx(bytes, page)
    return { content: imported.html, findings: imported.findings }
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
    void recordPreview(document.path, renderPreview(editor.state.doc, page))
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
    const created: CachedDoc = { entry: { status: 'loading' }, documentId: document.id, imageSources: new Map() }
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

async function encode(editor: Editor, format: string, imageSources: ReadonlyMap<string, string>): Promise<Uint8Array> {
  if (format === 'md' || format === 'markdown') {
    const { writeMarkdown } = await import('./io/markdown')
    return new TextEncoder().encode(writeMarkdown(editor.getJSON(), imageSources))
  }
  const { writeDocx } = await import('./io/docx-export')
  const { bytes, skippedImages } = await writeDocx(editor.getJSON(), page)
  if (skippedImages > 0) {
    toast.warning(`${skippedImages} ${skippedImages === 1 ? 'image was' : 'images were'} not saved`, {
      description: 'Word files can hold PNG, JPEG, GIF and BMP images that are part of the document, not linked from elsewhere.',
    })
  }
  return bytes
}

/**
 * Saves a document tab. A file whose import report is not lossless asks before it is overwritten,
 * since saving would change or lose what the report lists; "Save a copy" keeps the original intact.
 */
export async function saveDoc(documentId: string, saveAs: boolean): Promise<boolean> {
  const document = useDocumentsStore.getState().documents.find(candidate => candidate.id === documentId)
  const cached = cache.get(documentId)
  const editor = docEditor(documentId)
  if (document === undefined || cached === undefined || editor === undefined) return false

  let chooseNewPath = saveAs || document.path === undefined
  const imported = useFidelityStore.getState().reports[documentId]
  if (!chooseNewPath && imported !== undefined && imported.severity !== 'lossless') {
    const { value } = await askConfirm({
      title: 'Save over the original?',
      message: `Saving writes only what Sumi shows, so what the import report lists would change or be lost in ${imported.sourceName}. Saving a copy keeps the original file intact.`,
      actions: [{ label: 'Overwrite', value: 'overwrite' }, { label: 'Save a copy', value: 'copy', primary: true }],
    })
    if (value === 'cancel') return false
    chooseNewPath = value === 'copy'
  }

  const formats = saveFormats(document.extension)
  const suffix = chooseNewPath && !saveAs && document.path !== undefined ? ' (edited)' : ''
  const path = chooseNewPath
    ? await fileService.chooseSavePath({ defaultName: `${document.name}${suffix}.${formats[0] ?? 'docx'}`, extensions: formats })
    : document.path
  if (path === undefined) return false

  const format = extensionOf(path)
  const snapshot = editor.state.doc
  await fileService.write(path, await encode(editor, format, cached.imageSources))
  const file = await fileService.describe(path)
  cached.saved = snapshot

  const newId = useDocumentsStore.getState().setSaved(documentId, file)
  if (newId !== documentId) {
    cache.delete(documentId)
    cached.documentId = newId
    cache.set(newId, cached)
    emit()
  }
  // Edits made while the file was being written keep the tab dirty.
  useDocumentsStore.getState().setDirty(newId, !editor.state.doc.eq(snapshot))
  // The file on disk is now one Sumi wrote, so it holds nothing the editor does not show.
  useFidelityStore.getState().forget(documentId)
  if (imported !== undefined) useFidelityStore.getState().publish(newId, createImportReport(file.name, []))
  await recordRecent(file)
  await recordPreview(file.path, renderPreview(snapshot, page))
  toast.success(`Saved ${file.name}`)
  if (format !== 'docx') {
    toast.info('Markdown keeps text and structure only', { description: 'Underline, alignment and page breaks are saved in .docx.' })
  }
  return true
}
