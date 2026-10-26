import { BrowserWindow, dialog, shell, type IpcMainInvokeEvent } from 'electron'
import { basename } from 'node:path'
import type { CloseChoice, OpenFileOptions, SaveFileOptions } from '../../src/shared/shell'
import { NodeFileService } from './node-file-service'
import { handle } from './security'

const fileService = new NodeFileService()

const formatNames: Record<string, string> = {
  xlsx: 'Excel Workbook',
  csv: 'CSV (comma-separated values)',
  tsv: 'TSV (tab-separated values)',
  pdf: 'PDF document',
  docx: 'Word Document',
  md: 'Markdown',
  markdown: 'Markdown',
  pptx: 'PowerPoint Presentation',
}

function ownerOf(event: IpcMainInvokeEvent): BrowserWindow | undefined {
  return BrowserWindow.fromWebContents(event.sender) ?? undefined
}

/**
 * Registers the only main-process filesystem operations exposed to the renderer. `openInOtherWindow`
 * says whether a window other than the sender has a tab on a path; the save dialog refuses such a
 * path, since writing there would replace a file another window is editing.
 */
export function registerFileIpc(openInOtherWindow: (event: IpcMainInvokeEvent, path: string) => boolean): void {
  handle('file:show-open-dialog', async (event, options: OpenFileOptions) => {
    const properties: Array<'openFile' | 'multiSelections'> = ['openFile']
    if (options.allowMultiple === true) properties.push('multiSelections')

    const filters = options.extensions === undefined || options.extensions.length === 0
      ? undefined
      : [{ name: 'Supported files', extensions: [...options.extensions] }]

    const owner = ownerOf(event)
    const result = owner === undefined
      ? await dialog.showOpenDialog({ properties, filters })
      : await dialog.showOpenDialog(owner, { properties, filters })

    if (result.canceled) return []
    return Promise.all(result.filePaths.map(path => fileService.describeFile(path)))
  })

  handle('file:show-save-dialog', async (event, options: SaveFileOptions) => {
    const filters = options.extensions.map(extension => ({
      name: formatNames[extension] ?? extension.toUpperCase(),
      extensions: [extension],
    }))
    const owner = ownerOf(event)
    const settings = { defaultPath: options.defaultName, filters }
    const result = owner === undefined
      ? await dialog.showSaveDialog(settings)
      : await dialog.showSaveDialog(owner, settings)
    if (result.canceled) return undefined
    if (openInOtherWindow(event, result.filePath)) {
      throw new Error(`“${basename(result.filePath)}” is open in another window. Close it there first, or save with a different name.`)
    }
    return result.filePath
  })

  handle('file:confirm-close', async (event, documentName: string): Promise<CloseChoice> => {
    const settings = {
      type: 'warning' as const,
      message: `Do you want to save the changes you made to “${basename(documentName)}”?`,
      detail: 'Your changes will be lost if you don’t save them.',
      buttons: ['Save', 'Don’t Save', 'Cancel'],
      defaultId: 0,
      cancelId: 2,
    }
    const owner = ownerOf(event)
    const { response } = owner === undefined
      ? await dialog.showMessageBox(settings)
      : await dialog.showMessageBox(owner, settings)
    return response === 0 ? 'save' : response === 1 ? 'discard' : 'cancel'
  })

  handle('file:describe', (_event, path: string) => fileService.describeFile(path))
  handle('file:read', (_event, path: string) => fileService.readFile(path))
  handle('file:write', (_event, path: string, contents: Uint8Array) =>
    fileService.writeFile(path, contents),
  )
  handle('file:reveal', (_event, path: string) => shell.showItemInFolder(path))
}
