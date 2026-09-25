import type { ApplicationId } from './applications'
import type { CommandId } from './commands'
import type { Paper } from './print'

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

/**
 * The update offer as every window shows it. The main process keeps the release itself (its URLs
 * and files); the renderer only asks it to check, download or open the release notes.
 */
export type UpdateStatus =
  | { readonly state: 'none' }
  | { readonly state: 'available'; readonly version: string }
  | { readonly state: 'downloading'; readonly version: string; readonly percent: number }
  /** Verified and opened in Finder; installing is dragging the app into Applications. */
  | { readonly state: 'downloaded'; readonly version: string }
  | { readonly state: 'failed'; readonly version: string; readonly message: string }

/** The answer to Check for Updates. `available` also reaches every window as the offer. */
export type UpdateCheckResult =
  | { readonly outcome: 'available'; readonly version: string }
  | { readonly outcome: 'current'; readonly version: string }
  | { readonly outcome: 'failed'; readonly message: string }

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
  /** The paper printouts are laid out for: Letter in the US and a few other regions, A4 elsewhere. */
  printPaper(): Promise<Paper>
  /** Prints this window's page through the print dialog. Resolves false when cancelled. */
  print(): Promise<boolean>
  /** Prints this window's page to PDF bytes, with the page sizes its CSS sets. */
  printToPdf(): Promise<Uint8Array>
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
  checkForUpdates(): Promise<UpdateCheckResult>
  /** Downloads the offered release to Downloads, verifies it and opens it; progress arrives as status. */
  downloadUpdate(): Promise<void>
  openReleaseNotes(): Promise<void>
  onDocumentOpened(listener: (file: FileReference) => void): () => void
  onCommand(listener: (command: CommandId) => void): () => void
  /** Fires when the user closes the window or quits while documents are unsaved. */
  onCloseRequested(listener: () => void): () => void
  /** Receives the latest status at once, then each change. */
  onUpdateStatus(listener: (status: UpdateStatus) => void): () => void
  /** Fires once per check that finds a release, to show the offer as a toast. */
  onUpdateOffered(listener: () => void): () => void
}
