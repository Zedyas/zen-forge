import type { ApplicationId } from './applications'
import type { CommandId } from './commands'

export interface FileReference {
  readonly path: string
  readonly name: string
  readonly extension: string
  readonly size: number
  readonly lastModified: number
}

export interface OpenFileOptions {
  readonly extensions?: readonly string[]
  readonly allowMultiple?: boolean
}

export interface SaveFileOptions {
  readonly defaultName: string
  /** Offered formats; the first is selected. */
  readonly extensions: readonly string[]
}

export type Appearance = 'system' | 'light' | 'dark'
export type CloseChoice = 'save' | 'discard' | 'cancel'

/** Saved files open in a window, in tab order, and the active one; reopened on the next launch. */
export interface WindowSession {
  readonly paths: readonly string[]
  readonly activePath?: string
}

/** What a window reports to the main process: the title and proxy icon, the menu set, unsaved state, its session. */
export interface WindowState {
  readonly title: string
  /** The active tab's application (`home` for the Home tab); decides which menus the menu bar shows. */
  readonly active: ApplicationId
  readonly path?: string
  readonly edited: boolean
  readonly hasUnsaved: boolean
  readonly session: WindowSession
}

/** Per-window view toggles mirrored into the native View menu's check marks. */
export interface ViewState {
  readonly toolbar: boolean
  readonly inspector: boolean
}

export interface ShellBridge {
  readonly platform: string
  newWindow(): Promise<void>
  /** Opens a saved file as the only tab of a new window. */
  openInNewWindow(file: FileReference): Promise<void>
  /** Last session's tabs for this window, handed over once; undefined for windows that start empty. */
  takeSession(): Promise<WindowSession | undefined>
  showOpenDialog(options: OpenFileOptions): Promise<readonly FileReference[]>
  /** Resolves to the chosen path, or undefined when cancelled. */
  showSaveDialog(options: SaveFileOptions): Promise<string | undefined>
  confirmClose(documentName: string): Promise<CloseChoice>
  describeFile(path: string): Promise<FileReference>
  readFile(path: string): Promise<Uint8Array>
  writeFile(path: string, contents: Uint8Array): Promise<void>
  revealInFinder(path: string): Promise<void>
  pathForDroppedFile(file: File): string
  getAppearance(): Promise<Appearance>
  setAppearance(appearance: Appearance): Promise<void>
  setWindowState(state: WindowState): Promise<void>
  setViewState(state: ViewState): Promise<void>
  /**
   * Answers a close request: `true` closes the window (the renderer has saved or discarded every
   * unsaved document), `false` keeps it open and cancels a pending quit.
   */
  resolveClose(approved: boolean): Promise<void>
  onDocumentOpened(listener: (file: FileReference) => void): () => void
  onCommand(listener: (command: CommandId) => void): () => void
  /** Fires when the user closes the window or quits while documents are unsaved. */
  onCloseRequested(listener: () => void): () => void
}
