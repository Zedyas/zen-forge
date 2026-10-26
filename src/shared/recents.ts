import { isReadableExtension } from './applications'

/** How many recent files File > Open Recent and the command palette list. */
export const recentMenuLimit = 10

/** The most paths the main process keeps; Home lists no more than this many recent files. */
export const recentPathLimit = 40

const absolutePath = /^(\/|[A-Za-z]:[\\/]|\\\\)/

function isRecentPath(value: unknown): value is string {
  if (typeof value !== 'string' || value.length > 4096 || value.includes('\0') || !absolutePath.test(value)) return false
  const name = value.slice(Math.max(value.lastIndexOf('/'), value.lastIndexOf('\\')) + 1)
  return name.includes('.') && isReadableExtension(name.slice(name.lastIndexOf('.') + 1))
}

/**
 * The list the main process keeps after a window reports `payload`, which crosses IPC and so may be
 * anything. Entries that are not absolute paths to a file Zen Suzu opens are left out, and at most
 * `recentPathLimit` stay. A payload that is not a list, or that changes nothing, returns `current`
 * itself.
 */
export function acceptRecents(current: readonly string[], payload: unknown): readonly string[] {
  if (!Array.isArray(payload)) return current
  const next = payload.filter(isRecentPath).slice(0, recentPathLimit)
  return next.length === current.length && next.every((path, index) => path === current[index]) ? current : next
}

export interface RecentMenuItem {
  readonly path: string
  readonly label: string
}

function splitPath(path: string): { readonly name: string; readonly directory: string } {
  const end = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
  return { name: path.slice(end + 1), directory: path.slice(0, Math.max(end, 0)) }
}

/**
 * The files to list for a recents list ordered newest first: each path once, at most
 * `recentMenuLimit` of them. A file name shared by two listed files is followed by its folder,
 * or by its whole path when the folders share a name too.
 */
export function recentMenuItems(paths: readonly string[]): RecentMenuItem[] {
  const files = [...new Set(paths)].slice(0, recentMenuLimit).map(path => ({ path, ...splitPath(path) }))
  return files.map(file => {
    const folder = splitPath(file.directory).name
    const sameName = files.filter(other => other !== file && other.name === file.name)
    if (sameName.length === 0) return { path: file.path, label: file.name }
    const sameFolder = sameName.some(other => splitPath(other.directory).name === folder)
    return { path: file.path, label: `${file.name} — ${sameFolder ? file.directory : folder}` }
  })
}
