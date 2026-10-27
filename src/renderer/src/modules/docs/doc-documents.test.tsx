// @vitest-environment happy-dom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import type { Editor, JSONContent } from '@tiptap/core'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { FileReference } from '@shared/shell'

const { askConfirm, chooseDocumentPath, describeFile, writeFile } = vi.hoisted(() => ({
  askConfirm: vi.fn<(options: { readonly title: string }) => Promise<{ readonly value: string }>>(),
  chooseDocumentPath: vi.fn<() => Promise<string | undefined>>(),
  describeFile: vi.fn<(path: string) => Promise<FileReference>>(),
  writeFile: vi.fn<(path: string, bytes: Uint8Array) => Promise<void>>(),
}))
vi.mock('../../ui/ConfirmDialog', () => ({ askConfirm }))
vi.mock('../../app/document-actions', () => ({ chooseDocumentPath }))
vi.mock('../../services/file/IpcFileService', () => ({ fileService: { describe: describeFile, write: writeFile } }))
vi.mock('../../services/index/document-index', () => ({ recordPreview: async () => undefined, recordRecent: async () => undefined }))
vi.mock('sonner', () => ({ toast: { success: vi.fn(), info: vi.fn(), warning: vi.fn() } }))

import { useDocumentsStore } from '../../app/documents-store'
import { docEditor, releaseDoc, saveDoc, useDocEntry } from './doc-documents'

Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })

const saved: FileReference = { path: '/out/copy.docx', name: 'copy', extension: 'docx', size: 1, lastModified: 0 }
const line = (text: string): JSONContent => ({ type: 'paragraph', content: [{ type: 'text', text }] })
/** A Word file holds only pictures stored in the document, so this one comes back without its picture. */
const linkedPicture: JSONContent = { type: 'paragraph', content: [{ type: 'image', attrs: { src: 'https://example.com/a.png' } }] }

const cleanups: Array<() => void> = []
afterEach(() => cleanups.splice(0).forEach(cleanup => cleanup()))

beforeEach(() => {
  vi.clearAllMocks()
  chooseDocumentPath.mockResolvedValue(saved.path)
  describeFile.mockResolvedValue(saved)
  writeFile.mockResolvedValue(undefined)
})

/** An untitled document open in the store, with `content` in its editor. */
function openWith(content: JSONContent): { readonly id: string; readonly editor: Editor } {
  const id = useDocumentsStore.getState().openUntitled('docs')
  const document = useDocumentsStore.getState().documents.find(candidate => candidate.id === id)
  if (document === undefined) throw new Error('The document did not open.')
  const Host = (): null => {
    useDocEntry(document)
    return null
  }
  const host = window.document.createElement('div')
  const root = createRoot(host)
  act(() => root.render(<Host />))
  cleanups.push(() => {
    act(() => root.unmount())
    releaseDoc(useDocumentsStore.getState().documents.find(candidate => candidate.path === saved.path)?.id ?? id)
  })
  const editor = docEditor(id)
  if (editor === undefined) throw new Error('The editor did not load.')
  act(() => { editor.commands.setContent({ type: 'doc', content: [content] }) })
  return { id, editor }
}

describe('Save As of a new Word file', () => {
  it('writes without asking when the file reads back as the document shown', async () => {
    const { id } = openWith(line('Plain text'))
    await expect(saveDoc(id, true)).resolves.toBe(true)
    expect(askConfirm).not.toHaveBeenCalled()
    expect(writeFile).toHaveBeenCalledOnce()
  })

  it('asks before writing a file that would hold less, and writes nothing on Cancel', async () => {
    askConfirm.mockResolvedValue({ value: 'cancel' })
    const { id } = openWith(linkedPicture)
    await expect(saveDoc(id, true)).resolves.toBe(false)
    expect(askConfirm).toHaveBeenCalledOnce()
    expect(askConfirm.mock.calls[0]?.[0].title).toBe('Save with changes?')
    expect(writeFile).not.toHaveBeenCalled()
  })

  it('writes on Save, and keeps the tab unsaved when the document was edited during the save', async () => {
    askConfirm.mockResolvedValue({ value: 'save' })
    const { id, editor } = openWith(linkedPicture)
    writeFile.mockImplementation(async () => {
      act(() => { editor.commands.insertContentAt(editor.state.doc.content.size, line('Typed while saving')) })
    })
    await expect(saveDoc(id, true)).resolves.toBe(true)
    expect(writeFile).toHaveBeenCalledOnce()
    expect(useDocumentsStore.getState().documents.find(document => document.path === saved.path)?.dirty).toBe(true)
  })
})
