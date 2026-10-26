import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { FileReference } from '@shared/shell'

const { describeFile, forgetRecent, recordRecent } = vi.hoisted(() => ({
  describeFile: vi.fn<(path: string) => Promise<FileReference>>(),
  forgetRecent: vi.fn<(path: string) => Promise<void>>(async () => undefined),
  recordRecent: vi.fn<(file: FileReference) => Promise<void>>(async () => undefined),
}))
vi.mock('../services/file/IpcFileService', () => ({ fileService: { describe: describeFile } }))
vi.mock('../services/index/document-index', () => ({ forgetRecent, recordRecent }))
vi.mock('../services/platform/client', () => ({ platformClient: {} }))
vi.mock('../ui/ConfirmDialog', () => ({ askConfirm: async () => ({ value: 'cancel' }) }))
vi.mock('./editors', () => ({ editorFor: () => undefined }))

import { openRecent } from './document-actions'
import { homeTabId, useDocumentsStore } from './documents-store'

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

  it('drops a file that is gone from the list and reports why it did not open', async () => {
    describeFile.mockRejectedValue(new Error('“report.docx” was moved or deleted.'))
    await expect(openRecent(report.path)).rejects.toThrow('“report.docx” was moved or deleted.')
    expect(forgetRecent).toHaveBeenCalledWith(report.path)
    expect(useDocumentsStore.getState().documents).toEqual([])
  })
})
