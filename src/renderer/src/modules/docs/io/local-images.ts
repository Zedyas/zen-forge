/**
 * The file an image path in a Markdown file names, when it stays inside that file's folder:
 * `images/a.png` does, `/Users/…`, `~/…`, `file:…` and a `..` that climbs out of the folder do not.
 */
export function localImagePath(directory: string, src: string): string | undefined {
  let decoded = src
  try {
    decoded = decodeURI(src)
  } catch {
    // A literal % that is not an escape: the path is used as written.
  }
  if (/^([/~\\]|[a-z][a-z0-9+.-]*:)/i.test(decoded)) return undefined
  const parts: string[] = []
  for (const segment of decoded.split('/')) {
    if (segment === '' || segment === '.') continue
    if (segment !== '..') parts.push(segment)
    else if (parts.pop() === undefined) return undefined
  }
  return parts.length === 0 ? undefined : `${directory}${parts.join('/')}`
}
