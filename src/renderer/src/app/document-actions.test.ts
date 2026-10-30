import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { CloseChoice, FileReference } from '@shared/shell'
import type { DroppedFiles } from '../services/file/FileService'

const { describeFile, describeDroppedFiles, warn, forgetRecent, recordRecent, confirmClose, resolveClose, showInOtherWindow, editorFor } = vi.hoisted(() => ({
  describeFile: vi.fn<(path: string) => Promise<FileReference>>(),
  describeDroppedFiles: vi.fn<(files: ArrayLike<File>) => Promise<DroppedFiles>>(),
  warn: vi.fn<(title: string, options: { description: string }) => void>(),
  forgetRecent: vi.fn<(path: string) => Promise<void>>(async () => undefined),
  recordRecent: vi.fn<(file: FileReference) => Promise<void>>(async () => undefined),
  confirmClose: vi.fn<(name: string) => Promise<CloseChoice>>(),
  resolveClose: vi.fn<(approved: boolean) => Promise<void>>(async () => undefined),
  showInOtherWindow: vi.fn<(path: string) => Promise<boolean>>(async () => false),
  editorFor: vi.fn<(kind: string) => EditorHandler | undefined>(() => undefined),
}))
vi.mock('sonner', () => ({ toast: { warning: warn } }))
vi.mock('../services/file/IpcFileService', () => ({ fileService: { describe: describeFile, describeDroppedFiles } }))
vi.mock('../services/index/document-index', () => ({ forgetRecent, recordRecent }))
vi.mock('../services/platform/client', () => ({ platformClient: { confirmClose, resolveClose, showInOtherWindow } }))
vi.mock('../ui/ConfirmDialog', () => ({ askConfirm: async () => ({ value: 'cancel' }) }))
vi.mock('./editors', () => ({ editorFor }))

import { closeDocument, openDroppedFiles, openRecent, resolveWindowClose } from './document-actions'
import type { EditorHandler } from './editors'
import { homeTabId, useDocumentsStore, type OpenDocument } from './documents-store'

const report: FileReference = { path: '/files/report.docx', name: 'report', extension: 'docx', size: 1, lastModified: 0 }

describe('opening a recent file', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    useDocumentsStore.setState({ documents: [], activeId: homeTabId })
  })

  it('opens a file that is still there as a tab and keeps it in the list', async () => {
    describeFile.mockResolvedValue(report)
    await openRecent(report.path)
    expect(useDocumentsStore.getState().documents.map(document => document.id)).toEqual([report.path])
    expect(recordRecent).toHaveBeenCalledWith(report)
    expect(forgetRecent).not.toHaveBeenCalled()
  })

  it('shows a file another window has open in that window instead of opening a second tab', async () => {
    describeFile.mockResolvedValue(report)
    showInOtherWindow.mockResolvedValueOnce(true)
    await openRecent(report.path)
    expect(showInOtherWindow).toHaveBeenCalledWith(report.path)
    expect(useDocumentsStore.getState().documents).toEqual([])
  })

  it('drops a file that is gone from the list and reports why it did not open', async () => {
    describeFile.mockRejectedValue(new Error('“report.docx” was moved or deleted.'))
    await expect(openRecent(report.path)).rejects.toThrow('“report.docx” was moved or deleted.')
    expect(forgetRecent).toHaveBeenCalledWith(report.path)
    expect(useDocumentsStore.getState().documents).toEqual([])
  })
})

describe('dropping files', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    useDocumentsStore.setState({ documents: [], activeId: homeTabId })
  })

  it('opens the files dropped with a folder and names the folder it skipped', async () => {
    describeDroppedFiles.mockResolvedValue({ files: [report], skipped: [{ name: 'Invoices', reason: 'The selected path is not a file.' }] })
    await openDroppedFiles([])
    expect(useDocumentsStore.getState().documents.map(document => document.id)).toEqual([report.path])
    expect(warn).toHaveBeenCalledWith('Skipped “Invoices”', { description: 'The selected path is not a file.' })
  })
})

describe('closing the window', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    useDocumentsStore.setState({ documents: [{ id: report.path, name: 'report', extension: 'docx', kind: 'docs', path: report.path, dirty: true }], activeId: homeTabId })
  })

  it('asks once about each unsaved document when a second close request arrives while it is asking', async () => {
    let answer: (choice: CloseChoice) => void = () => undefined
    confirmClose.mockReturnValue(new Promise(resolve => { answer = resolve }))
    // A quit re-sends the close request to every window still open whenever another window closes.
    const first = resolveWindowClose()
    const second = resolveWindowClose()
    await vi.waitFor(() => expect(confirmClose).toHaveBeenCalledTimes(1))
    answer('discard')
    await Promise.all([first, second])
    expect(confirmClose).toHaveBeenCalledTimes(1)
    expect(resolveClose.mock.calls).toEqual([[true]])
  })
})

describe('closing a tab', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    useDocumentsStore.setState({ documents: [{ id: 'untitled:0', name: 'Untitled', extension: 'docx', kind: 'docs', dirty: true }], activeId: 'untitled:0' })
  })

  it('closes a new document once Save has written it, though saving gave the tab its file path', async () => {
    const release = vi.fn<(id: string) => void>()
    editorFor.mockReturnValue({
      run: () => undefined,
      release,
      // As the editors do: a first save re-keys the tab to the path chosen in the save dialog.
      save: async id => {
        useDocumentsStore.getState().setSaved(id, report)
        return true
      },
    })
    confirmClose.mockResolvedValue('save')
    expect(await closeDocument('untitled:0')).toBe(true)
    expect(useDocumentsStore.getState().documents).toEqual([])
    expect(release).toHaveBeenCalledWith(report.path)
  })

  it('closes the saved tab, not one opened while Save was writing it', async () => {
    const existing: OpenDocument = { id: '/files/budget.xlsx', name: 'budget', extension: 'xlsx', kind: 'sheets', path: '/files/budget.xlsx', dirty: false }
    useDocumentsStore.setState(state => ({ documents: [existing, ...state.documents] }))
    let opened = ''
    editorFor.mockReturnValue({
      run: () => undefined,
      release: () => undefined,
      save: async id => {
        // A new spreadsheet joins the spreadsheets, ahead of the document being saved.
        opened = useDocumentsStore.getState().openUntitled('sheets')
        useDocumentsStore.getState().setDirty(opened, true)
        useDocumentsStore.getState().setSaved(id, report)
        return true
      },
    })
    confirmClose.mockResolvedValue('save')
    expect(await closeDocument('untitled:0')).toBe(true)
    expect(useDocumentsStore.getState().documents.map(document => document.id)).toEqual([existing.id, opened])
  })

  it('asks again instead of closing when typing arrived while Save was writing the file', async () => {
    const release = vi.fn<(id: string) => void>()
    editorFor.mockReturnValue({
      run: () => undefined,
      release,
      // As the editors do: edits made while the file was written keep the tab unsaved.
      save: async id => {
        const saved = useDocumentsStore.getState().setSaved(id, report)
        useDocumentsStore.getState().setDirty(saved, true)
        return true
      },
    })
    confirmClose.mockResolvedValueOnce('save').mockResolvedValueOnce('cancel')
    expect(await closeDocument('untitled:0')).toBe(false)
    expect(confirmClose).toHaveBeenCalledTimes(2)
    expect(release).not.toHaveBeenCalled()
    expect(useDocumentsStore.getState().documents).toMatchObject([{ id: report.path, dirty: true }])
  })
})
