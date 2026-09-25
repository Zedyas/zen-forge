import { contextBridge, ipcRenderer, webUtils } from 'electron'
import type { CommandId } from '../../src/shared/commands'
import type { FileReference, ShellBridge } from '../../src/shared/shell'

const documentListeners = new Set<(file: FileReference) => void>()
const pendingDocuments: FileReference[] = []
const commandListeners = new Set<(command: CommandId) => void>()
const pendingCommands: CommandId[] = []
const closeListeners = new Set<() => void>()

// Documents can arrive before React subscribes (a Finder open launches the window); hold them until then.
ipcRenderer.on('shell:document-opened', (_event, file: FileReference) => {
  if (documentListeners.size === 0) {
    pendingDocuments.push(file)
    return
  }
  documentListeners.forEach(listener => listener(file))
})

ipcRenderer.on('shell:command', (_event, command: CommandId) => {
  if (commandListeners.size === 0) {
    pendingCommands.push(command)
    return
  }
  commandListeners.forEach(listener => listener(command))
})

ipcRenderer.on('shell:close-requested', () => {
  // With nothing listening there is no unsaved state to resolve, so the window may close.
  if (closeListeners.size === 0) void ipcRenderer.invoke('shell:resolve-close', true)
  closeListeners.forEach(listener => listener())
})

const shellBridge: ShellBridge = {
  platform: process.platform,
  newWindow: () => ipcRenderer.invoke('shell:new-window'),
  openInNewWindow: file => ipcRenderer.invoke('shell:open-in-new-window', file),
  takeSession: () => ipcRenderer.invoke('shell:take-session'),
  showOpenDialog: options => ipcRenderer.invoke('file:show-open-dialog', options),
  showSaveDialog: options => ipcRenderer.invoke('file:show-save-dialog', options),
  confirmClose: documentName => ipcRenderer.invoke('file:confirm-close', documentName),
  describeFile: path => ipcRenderer.invoke('file:describe', path),
  readFile: path => ipcRenderer.invoke('file:read', path),
  writeFile: (path, contents) => ipcRenderer.invoke('file:write', path, contents),
  revealInFinder: path => ipcRenderer.invoke('file:reveal', path),
  pathForDroppedFile: file => webUtils.getPathForFile(file),
  getAppearance: () => ipcRenderer.invoke('shell:get-appearance'),
  setAppearance: appearance => ipcRenderer.invoke('shell:set-appearance', appearance),
  setWindowState: state => ipcRenderer.invoke('shell:set-window-state', state),
  setViewState: state => ipcRenderer.invoke('shell:set-view-state', state),
  resolveClose: approved => ipcRenderer.invoke('shell:resolve-close', approved),
  onDocumentOpened: listener => {
    documentListeners.add(listener)
    pendingDocuments.splice(0).forEach(file => listener(file))
    return () => documentListeners.delete(listener)
  },
  onCommand: listener => {
    commandListeners.add(listener)
    pendingCommands.splice(0).forEach(command => listener(command))
    return () => commandListeners.delete(listener)
  },
  onCloseRequested: listener => {
    closeListeners.add(listener)
    return () => closeListeners.delete(listener)
  },
}

contextBridge.exposeInMainWorld('desktop', shellBridge)
