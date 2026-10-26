import { toast } from 'sonner'
import { applicationForExtension, applications, findApplication, isWritableExtension, type EditorApplicationId } from '@shared/applications'
import type { FileReference, SaveFileOptions } from '@shared/shell'
import { fileService } from '../services/file/IpcFileService'
import { recordRecent } from '../services/index/document-index'
import { platformClient } from '../services/platform/client'
import { askConfirm } from '../ui/ConfirmDialog'
import { editorFor } from './editors'
import { changedOnDisk, openInOtherTab, useDocumentsStore, type OpenDocument } from './documents-store'

function findDocument(id: string): OpenDocument | undefined {
  return useDocumentsStore.getState().documents.find(candidate => candidate.id === id)
}

function openableExtensions(kind?: EditorApplicationId): readonly string[] {
  const candidates = kind === undefined ? applications : [findApplication(kind)]
  return candidates.filter(candidate => candidate.available).flatMap(candidate => candidate.opens)
}

function isOpenable(file: FileReference): boolean {
  const kind = applicationForExtension(file.extension)
  return kind !== undefined && findApplication(kind).available
}

/** Opens files as tabs in this window and records them in recents. */
export async function openFiles(files: readonly FileReference[]): Promise<void> {
  for (const file of files) {
    if (!isOpenable(file)) throw new Error(`.${file.extension || 'unknown'} files cannot be opened yet.`)
    useDocumentsStore.getState().openFile(file)
    await recordRecent(file)
  }
}

/** Shows the open dialog, limited to one application's file types when `kind` is given. */
export async function chooseAndOpenFiles(kind?: EditorApplicationId): Promise<void> {
  await openFiles(await fileService.chooseFiles({ extensions: openableExtensions(kind), allowMultiple: true }))
}

export async function openDroppedFiles(files: FileList): Promise<void> {
  const openable = (await fileService.describeDroppedFiles(files)).filter(isOpenable)
  if (openable.length === 0) {
    throw new Error(`Drop a ${openableExtensions().map(extension => `.${extension}`).join(', ')} file.`)
  }
  await openFiles(openable)
}

/**
 * Asks where to save a document. A file open in another tab is refused: this tab would take that
 * tab's place, and the other tab's unsaved changes would be lost. The main process refuses files
 * open in other windows the same way.
 */
export async function chooseDocumentPath(id: string, options: SaveFileOptions): Promise<string | undefined> {
  const path = await fileService.chooseSavePath(options)
  const other = path === undefined ? undefined : openInOtherTab(useDocumentsStore.getState(), id, path)
  if (other !== undefined) {
    throw new Error(`“${other.name}” is open in another tab. Close that tab first, or save with a different name.`)
  }
  return path
}

/**
 * Before saving over a file that another app changed since this tab opened or saved it, asks:
 * Overwrite, or Save a copy, which keeps the other app's version. Resolves to the choice or `cancel`.
 */
async function confirmChangedOnDisk(document: OpenDocument): Promise<string> {
  if (document.path === undefined || document.onDisk === undefined || !isWritableExtension(document.extension)) return 'overwrite'
  // A file deleted since is simply written again.
  const current = await fileService.describe(document.path).catch(() => undefined)
  if (current === undefined || !changedOnDisk(document.onDisk, current)) return 'overwrite'
  const { value } = await askConfirm({
    title: 'Save over the changed file?',
    message: `Another app changed ${current.name} after Zen Suzu opened or last saved it. Overwriting replaces those changes. Saving a copy keeps both versions.`,
    actions: [{ label: 'Overwrite', value: 'overwrite' }, { label: 'Save a copy', value: 'copy', primary: true }],
  })
  return value
}

/** Saves one document through its module; reports failure as a toast and resolves false. */
export async function saveDocument(id: string, saveAs: boolean): Promise<boolean> {
  const document = findDocument(id)
  const editor = document === undefined ? undefined : editorFor(document.kind)
  if (document === undefined || editor === undefined) return false
  try {
    const choice = saveAs ? 'overwrite' : await confirmChangedOnDisk(document)
    if (choice === 'cancel') return false
    return await editor.save(id, saveAs || choice === 'copy')
  } catch (error) {
    toast.error('Could not save', { description: error instanceof Error ? error.message : 'The file could not be written.' })
    return false
  }
}

/** Writes whatever is still being typed in a document's editor, so its unsaved state is up to date. */
function flushTyping(id: string): void {
  const document = findDocument(id)
  if (document !== undefined) editorFor(document.kind)?.flush?.(id)
}

/** Closes a tab, asking Save / Don't Save / Cancel first when it has unsaved changes. */
export async function closeDocument(id: string): Promise<boolean> {
  flushTyping(id)
  const document = findDocument(id)
  if (document === undefined) return true
  if (document.dirty) {
    useDocumentsStore.getState().select(id)
    const choice = await platformClient.confirmClose(document.name)
    if (choice === 'cancel') return false
    if (choice === 'save' && !(await saveDocument(id, false))) return false
  }
  editorFor(document.kind)?.release(id)
  useDocumentsStore.getState().remove(id)
  return true
}

/**
 * Moves a saved tab into a new window. Its unsaved state lives in this window's memory, so a
 * document with changes has to be saved first.
 */
export async function moveToNewWindow(id: string): Promise<void> {
  flushTyping(id)
  const document = findDocument(id)
  if (document === undefined) return
  if (document.path === undefined || document.dirty) {
    throw new Error(`Save ${document.name} first, then move it to a new window.`)
  }
  await platformClient.openInNewWindow(await fileService.describe(document.path))
  editorFor(document.kind)?.release(id)
  useDocumentsStore.getState().remove(id)
}

/** Resolves every unsaved document before the window closes or the app quits. */
export async function resolveWindowClose(): Promise<void> {
  useDocumentsStore.getState().documents.forEach(document => flushTyping(document.id))
  const unsaved = useDocumentsStore.getState().documents.filter(document => document.dirty)
  for (const document of unsaved) {
    if (!(await closeDocument(document.id))) {
      await platformClient.resolveClose(false)
      return
    }
  }
  await platformClient.resolveClose(true)
}
