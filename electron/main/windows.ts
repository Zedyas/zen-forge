/*
 * The side-effect-free decisions about windows: which one a file goes to, whether a close waits for
 * the page, and the Edit menu's Undo and Redo. `index.ts` creates the windows, menus and messages.
 */

import type { MenuItemConstructorOptions, WebContents } from 'electron'
import type { WindowState } from '../../src/shared/shell'

/** What the decisions read from a window. */
export interface WindowFacts {
  readonly window: { readonly webContents: Pick<WebContents, 'isCrashed'> }
  readonly state: Pick<WindowState, 'session'>
}

/**
 * A window other than `sender`'s with a tab on `path`. A window whose page crashed lost its tabs,
 * so it holds no file: sending one there would show nothing, and saving to it elsewhere is allowed.
 */
export function windowWith<Window extends WindowFacts>(windows: readonly Window[], path: string, sender?: unknown): Window | undefined {
  return windows.find(record =>
    record.window.webContents !== sender && !record.window.webContents.isCrashed() && record.state.session.paths.includes(path))
}

/**
 * The window a file or command from outside goes to: the window that has the file, else `current`,
 * else the newest window whose page still runs. Undefined means a new window.
 */
export function targetFor<Window extends WindowFacts>(windows: readonly Window[], current: Window | undefined, path?: string): Window | undefined {
  const holder = path === undefined ? undefined : windowWith(windows, path)
  if (holder !== undefined) return holder
  if (current !== undefined && !current.window.webContents.isCrashed()) return current
  return windows.findLast(record => !record.window.webContents.isCrashed())
}

/** What a close reads from the window that is closing. */
export interface ClosingWindow {
  readonly window: { readonly webContents: Pick<WebContents, 'isCrashed'> }
  readonly state: Pick<WindowState, 'hasUnsaved'>
  readonly closeApproved: boolean
}

/**
 * True when a close has to wait for the page to ask about unsaved documents. A crashed page cannot
 * answer, and its unsaved changes went with it; waiting would keep the window, and a quit, stuck.
 */
export function closeWaitsForPage(record: ClosingWindow): boolean {
  return record.state.hasUnsaved && !record.closeApproved && !record.window.webContents.isCrashed()
}

/**
 * Undo and Redo for the Edit menu: the active document's. Home has no document to undo; its text
 * fields still need Undo and Redo, which macOS reaches only through the menu.
 */
export function historyItems(edit: readonly MenuItemConstructorOptions[]): MenuItemConstructorOptions[] {
  const history = edit.filter(item => item.id === 'edit.undo' || item.id === 'edit.redo')
  return history.length > 0 ? history : [{ role: 'undo' }, { role: 'redo' }]
}
