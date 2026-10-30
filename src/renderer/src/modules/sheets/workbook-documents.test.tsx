// @vitest-environment happy-dom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it, vi } from 'vitest'
import type { FileReference } from '@shared/shell'

const { writeFile } = vi.hoisted(() => ({ writeFile: vi.fn<(path: string, bytes: Uint8Array) => Promise<void>>() }))
const saved: FileReference = { path: '/out/book.xlsx', name: 'book.xlsx', extension: 'xlsx', size: 1, lastModified: 0 }
vi.mock('../../ui/ConfirmDialog', () => ({ askConfirm: vi.fn() }))
vi.mock('../../app/document-actions', () => ({ chooseDocumentPath: async () => saved.path, openFiles: vi.fn() }))
vi.mock('../../services/file/IpcFileService', () => ({ fileService: { describe: async () => saved, write: writeFile } }))
vi.mock('../../services/index/document-index', () => ({ recordPreview: async () => undefined, recordRecent: async () => undefined }))
vi.mock('../../services/print/print', () => ({ openPrintDialog: vi.fn(), printToPdf: vi.fn() }))
vi.mock('sonner', () => ({ toast: { success: vi.fn(), info: vi.fn() } }))

import { useDocumentsStore } from '../../app/documents-store'
import { releaseWorkbook, saveWorkbook, useWorkbookEntry, type WorkbookEntry } from './workbook-documents'

Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })

it('keeps a spreadsheet tab unsaved when it was edited while Save wrote the file', async () => {
  const id = useDocumentsStore.getState().openUntitled('sheets')
  const document = useDocumentsStore.getState().documents.find(candidate => candidate.id === id)
  if (document === undefined) throw new Error('The spreadsheet did not open.')
  const seen: { entry?: WorkbookEntry } = {}
  const Host = (): null => {
    seen.entry = useWorkbookEntry(document)
    return null
  }
  const root = createRoot(window.document.createElement('div'))
  act(() => root.render(<Host />))
  const { entry } = seen
  if (entry?.status !== 'ready') throw new Error('The workbook did not load.')
  const { workbook } = entry
  const sheetId = workbook.sheets()[0]?.id ?? 0
  workbook.setCellInput(sheetId, [0, 0], 1)
  writeFile.mockImplementation(async () => {
    workbook.setCellInput(sheetId, [0, 1], 2)
  })

  await act(async () => {
    expect(await saveWorkbook(id, false)).toBe(true)
  })

  expect(useDocumentsStore.getState().documents.find(candidate => candidate.path === saved.path)?.dirty).toBe(true)
  act(() => root.unmount())
  releaseWorkbook(saved.path)
})
