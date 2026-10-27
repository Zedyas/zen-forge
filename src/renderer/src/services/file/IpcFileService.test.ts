import { afterEach, describe, expect, it, vi } from 'vitest'
import type { FileReference, ShellBridge } from '@shared/shell'
import { IpcFileService } from './IpcFileService'

const report: FileReference = { path: '/files/report.docx', name: 'report.docx', extension: 'docx', size: 1, lastModified: 0 }

describe('dropped files', () => {
  afterEach(() => vi.unstubAllGlobals())

  it('describes the files and skips a folder dropped with them, saying why', async () => {
    const bridge: Pick<ShellBridge, 'describeFile' | 'pathForDroppedFile'> = {
      pathForDroppedFile: file => `/files/${file.name}`,
      describeFile: async path => {
        if (path === report.path) return report
        throw new Error("Error invoking remote method 'file:describe': Error: The selected path is not a file.")
      },
    }
    vi.stubGlobal('window', { desktop: bridge })
    const dropped = await new IpcFileService().describeDroppedFiles([new File([], 'report.docx'), new File([], 'Invoices')])
    expect(dropped).toEqual({ files: [report], skipped: [{ name: 'Invoices', reason: 'The selected path is not a file.' }] })
  })
})
