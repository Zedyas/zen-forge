/** How many recent files File > Open Recent and the command palette list. */
export const recentMenuLimit = 10

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
