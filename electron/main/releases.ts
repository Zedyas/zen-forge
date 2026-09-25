/*
 * The side-effect-free half of the updater: reading GitHub's release list, ordering versions and
 * reading SHA256SUMS.txt. `updates.ts` does the network, files and windows.
 */

const repository = 'https://github.com/Zedyas/zen-forge'
export const releasesApiUrl = 'https://api.github.com/repos/Zedyas/zen-forge/releases?per_page=20'
const releasePagePrefix = `${repository}/releases/`
const downloadPrefix = `${repository}/releases/download/`

/** A release newer than the running app. Its URLs are already checked to point into Zendo's releases. */
export interface Release {
  /** `0.2.0`, without the tag's `v`. */
  readonly version: string
  readonly pageUrl: string
  readonly dmg: { readonly name: string; readonly url: string; readonly size: number }
  readonly checksumsUrl: string
}

interface Version {
  readonly text: string
  readonly core: readonly [number, number, number]
  /** `beta.2` becomes `['beta', '2']`; empty for a release. */
  readonly prerelease: readonly string[]
}

/** `1.2.3`, `v1.2.3` or `1.2.3-beta.2`; build metadata (`+…`) is allowed and ignored. */
export function parseVersion(text: string): Version | undefined {
  const match = /^v?(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?(?:\+[0-9A-Za-z.-]+)?$/.exec(text)
  if (match === null) return undefined
  const [, major = '', minor = '', patch = '', prerelease] = match
  return {
    text: text.replace(/^v/, ''),
    core: [Number(major), Number(minor), Number(patch)],
    prerelease: prerelease === undefined ? [] : prerelease.split('.'),
  }
}

/**
 * Negative when `a` is older, as semver.org §11 orders versions: numbers first, then a pre-release
 * sorts before its release, and `beta.2` before `beta.10`.
 */
export function compareVersions(a: Version, b: Version): number {
  for (let index = 0; index < 3; index += 1) {
    if (a.core[index] !== b.core[index]) return a.core[index] - b.core[index]
  }
  if (a.prerelease.length === 0 || b.prerelease.length === 0) return b.prerelease.length - a.prerelease.length
  for (let index = 0; index < Math.min(a.prerelease.length, b.prerelease.length); index += 1) {
    const left = a.prerelease[index]
    const right = b.prerelease[index]
    if (left === right) continue
    const leftNumeric = /^\d+$/.test(left)
    const rightNumeric = /^\d+$/.test(right)
    if (leftNumeric && rightNumeric) return Number(left) - Number(right)
    if (leftNumeric !== rightNumeric) return leftNumeric ? -1 : 1
    return left < right ? -1 : 1
  }
  return a.prerelease.length - b.prerelease.length
}

/** A JSON object's fields; anything else has none. GitHub's answer is network data, so each field is checked. */
function fields(value: unknown): ReadonlyMap<string, unknown> {
  return typeof value === 'object' && value !== null ? new Map(Object.entries(value)) : new Map()
}

/** The normalized URL when it lies under `prefix`, so `..` segments cannot step out of it. */
function urlUnder(value: unknown, prefix: string): string | undefined {
  if (typeof value !== 'string' || !URL.canParse(value)) return undefined
  const { href } = new URL(value)
  return href.startsWith(prefix) ? href : undefined
}

function findAsset(assets: unknown, name: string): Release['dmg'] | undefined {
  if (!Array.isArray(assets)) return undefined
  for (const asset of assets.map(fields)) {
    const url = urlUnder(asset.get('browser_download_url'), downloadPrefix)
    const size = asset.get('size')
    if (asset.get('name') === name && url !== undefined && typeof size === 'number') return { name, url, size }
  }
  return undefined
}

/**
 * The highest release above `currentVersion` that has both `Zendo-<version>-<arch>.dmg` and
 * SHA256SUMS.txt. Drafts never count. Pre-releases count only while no full release exists,
 * which is the case until builds are notarized.
 */
export function selectRelease(body: unknown, currentVersion: string, arch: string): Release | undefined {
  if (!Array.isArray(body)) throw new Error('GitHub answered with something other than a release list.')
  const current = parseVersion(currentVersion)
  if (current === undefined) throw new Error(`${currentVersion} is not a version.`)

  const published = body.map(fields).filter(release => release.get('draft') === false)
  const full = published.filter(release => release.get('prerelease') === false)
  let best: { version: Version; release: Release } | undefined
  for (const release of full.length > 0 ? full : published) {
    const tag = release.get('tag_name')
    const version = typeof tag === 'string' ? parseVersion(tag) : undefined
    if (version === undefined || compareVersions(version, current) <= 0) continue
    if (best !== undefined && compareVersions(version, best.version) <= 0) continue
    const pageUrl = urlUnder(release.get('html_url'), releasePagePrefix)
    const dmg = findAsset(release.get('assets'), `Zendo-${version.text}-${arch}.dmg`)
    const checksums = findAsset(release.get('assets'), 'SHA256SUMS.txt')
    if (pageUrl === undefined || dmg === undefined || checksums === undefined) continue
    best = { version, release: { version: version.text, pageUrl, dmg, checksumsUrl: checksums.url } }
  }
  return best?.release
}

/** The lowercase SHA-256 that `shasum -a 256` output lists for `fileName`, or undefined when it lists none. */
export function checksumFor(sums: string, fileName: string): string | undefined {
  for (const line of sums.split(/\r?\n/)) {
    // `<64 hex>  <name>`; binary mode writes `<64 hex> *<name>`.
    const match = /^([0-9a-fA-F]{64}) [ *](.+)$/.exec(line)
    if (match?.[2] === fileName) return match[1]?.toLowerCase()
  }
  return undefined
}
