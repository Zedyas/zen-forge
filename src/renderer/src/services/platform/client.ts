import type { CommandId } from '@shared/commands'
import type { Paper } from '@shared/print'
import type { Appearance, CloseChoice, FileReference, UpdateCheckResult, UpdateStatus, ViewState, WindowSession, WindowState } from '@shared/shell'
import { getBridge, mainProcessMessage, requireBridge } from '../file/IpcFileService'

/** Keeps renderer components independent from the preload bridge and its browser fallback. */
export const platformClient = {
  async newWindow(): Promise<void> {
    await requireBridge('New windows').newWindow()
  },

  async openInNewWindow(file: FileReference): Promise<void> {
    await requireBridge('New windows').openInNewWindow(file)
  },

  async takeSession(): Promise<WindowSession | undefined> {
    return getBridge()?.takeSession()
  },

  confirmClose(documentName: string): Promise<CloseChoice> {
    const bridge = getBridge()
    if (bridge !== undefined) return bridge.confirmClose(documentName)
    return Promise.resolve(window.confirm(`Discard changes to ${documentName}?`) ? 'discard' : 'cancel')
  },

  async revealInFinder(path: string): Promise<void> {
    await getBridge()?.revealInFinder(path)
  },

  printPaper(): Promise<Paper> {
    return requireBridge('Printing').printPaper()
  },

  print(): Promise<boolean> {
    return mainProcessMessage(requireBridge('Printing').print())
  },

  printToPdf(): Promise<Uint8Array> {
    return mainProcessMessage(requireBridge('Exporting PDF').printToPdf())
  },

  async getAppearance(): Promise<Appearance> {
    return (await getBridge()?.getAppearance()) ?? 'system'
  },

  async setAppearance(appearance: Appearance): Promise<void> {
    await getBridge()?.setAppearance(appearance)
  },

  async setWindowState(state: WindowState): Promise<void> {
    await getBridge()?.setWindowState(state)
  },

  async setViewState(state: ViewState): Promise<void> {
    await getBridge()?.setViewState(state)
  },

  async resolveClose(approved: boolean): Promise<void> {
    await getBridge()?.resolveClose(approved)
  },

  checkForUpdates(): Promise<UpdateCheckResult> {
    return requireBridge('Checking for updates').checkForUpdates()
  },

  async downloadUpdate(): Promise<void> {
    await getBridge()?.downloadUpdate()
  },

  async openReleaseNotes(): Promise<void> {
    await getBridge()?.openReleaseNotes()
  },

  onDocumentOpened(listener: (file: FileReference) => void): () => void {
    return getBridge()?.onDocumentOpened(listener) ?? (() => undefined)
  },

  onCommand(listener: (command: CommandId) => void): () => void {
    return getBridge()?.onCommand(listener) ?? (() => undefined)
  },

  onCloseRequested(listener: () => void): () => void {
    return getBridge()?.onCloseRequested(listener) ?? (() => undefined)
  },

  onUpdateStatus(listener: (status: UpdateStatus) => void): () => void {
    return getBridge()?.onUpdateStatus(listener) ?? (() => undefined)
  },

  onUpdateOffered(listener: () => void): () => void {
    return getBridge()?.onUpdateOffered(listener) ?? (() => undefined)
  },
}
