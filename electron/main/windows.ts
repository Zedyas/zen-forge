/*
 * The side-effect-free decisions about windows: which one a file goes to. `index.ts` creates the
 * windows, focuses them and sends to them.
 */

import type { WebContents } from 'electron'
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
