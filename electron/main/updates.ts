import { app, net, shell } from 'electron'
import { createHash } from 'node:crypto'
import { rmSync } from 'node:fs'
import { link, open, rm, type FileHandle } from 'node:fs/promises'
import { basename, extname, join } from 'node:path'
import { suiteName } from '../../src/shared/applications'
import type { UpdateCheckResult, UpdateStatus } from '../../src/shared/shell'
import { checksumFor, releasesApiUrl, selectRelease, type Release } from './releases'
import { handle } from './security'
import { readSettings, setLastUpdateCheck } from './settings'

/*
 * Updates for an app without a Developer ID signature. Electron's autoUpdater (Squirrel.Mac) will
 * not install into an ad-hoc signed app, so Zendo downloads the release's .dmg, checks it against
 * the release's SHA256SUMS.txt and opens it; the user drags the new Zendo into Applications.
 */

const day = 24 * 60 * 60 * 1000
/** A download that receives nothing for this long is dropped, so the user can try again. */
const stallTimeout = 30_000

let broadcast: (channel: string, payload?: unknown) => void = () => undefined
/** The release on offer. Only the main process holds its URLs; the renderer never names a URL or path. */
let offered: Release | undefined
let status: UpdateStatus = { state: 'none' }
let checking: Promise<UpdateCheckResult> | undefined
/** The unverified download in progress, removed if Zendo quits before it finishes. */
let partInProgress: string | undefined

function setStatus(next: UpdateStatus): void {
  status = next
  broadcast('update:status', status)
}

/** For a window that opens after a check, so its Home shows the offer too. */
export function currentUpdateStatus(): UpdateStatus {
  return status
}

async function request(url: string, accept: string, signal: AbortSignal): Promise<Response> {
  const response = await net.fetch(url, { headers: { Accept: accept, 'User-Agent': `${suiteName}/${app.getVersion()}` }, signal })
  // Rate limiting (403 or 429) fails the same way as any other refusal.
  if (!response.ok) throw new Error(`GitHub answered ${response.status} for ${url}.`)
  return response
}

async function runCheck(): Promise<UpdateCheckResult> {
  let release: Release | undefined
  try {
    const response = await request(releasesApiUrl, 'application/vnd.github+json', AbortSignal.timeout(15_000))
    release = selectRelease(await response.json(), app.getVersion(), process.arch)
  } catch {
    return { outcome: 'failed', message: 'Couldn’t reach GitHub. Try again later.' }
  }
  setLastUpdateCheck(Date.now())
  // A running download keeps its release and progress.
  if (status.state !== 'downloading') {
    offered = release
    setStatus(release === undefined ? { state: 'none' } : { state: 'available', version: release.version })
  }
  if (release === undefined) return { outcome: 'current', version: app.getVersion() }
  broadcast('update:offered')
  return { outcome: 'available', version: release.version }
}

/** Asks GitHub for a newer release. Checks that overlap share one request. */
function checkForUpdates(): Promise<UpdateCheckResult> {
  checking ??= runCheck().finally(() => {
    checking = undefined
  })
  return checking
}

/**
 * Runs `claim` on `name` in `folder`, or on `name (1)`, `name (2)` and so on while it fails because the
 * name is taken, as browsers name repeat downloads. Resolves the path it claimed; never overwrites a file.
 */
async function claimName<Result>(folder: string, name: string, claim: (path: string) => Promise<Result>): Promise<{ path: string; result: Result }> {
  const extension = extname(name)
  for (let copy = 0; copy < 100; copy += 1) {
    const path = join(folder, copy === 0 ? name : `${basename(name, extension)} (${copy})${extension}`)
    try {
      return { path, result: await claim(path) }
    } catch (error) {
      if (!(error instanceof Error && 'code' in error && error.code === 'EEXIST')) throw error
    }
  }
  throw new Error(`Too many copies of ${name} in ${folder}.`)
}

/** Streams the dmg into `file`, reporting whole percents, and returns its SHA-256 in hex. */
async function save(dmg: Release['dmg'], file: FileHandle, onPercent: (percent: number) => void): Promise<string> {
  const stalled = new AbortController()
  const timer = setTimeout(() => stalled.abort(), stallTimeout)
  try {
    const response = await request(dmg.url, 'application/octet-stream', stalled.signal)
    if (response.body === null) throw new Error(`${dmg.name} arrived empty.`)
    const reader = response.body.getReader()
    const hash = createHash('sha256')
    let received = 0
    let percent = 0
    for (let chunk = await reader.read(); !chunk.done; chunk = await reader.read()) {
      timer.refresh()
      hash.update(chunk.value)
      await file.write(chunk.value)
      received += chunk.value.byteLength
      if (received > dmg.size) throw new Error(`${dmg.name} is larger than GitHub listed.`)
      const next = Math.min(100, Math.floor((received / dmg.size) * 100))
      if (next !== percent) {
        percent = next
        onPercent(percent)
      }
    }
    return hash.digest('hex')
  } finally {
    clearTimeout(timer)
    await file.close()
  }
}

/**
 * Downloads the offered dmg to Downloads as a `.part` file, checks its SHA-256 against the release's
 * SHA256SUMS.txt, and only then gives it its .dmg name and opens it. A download cut short, by a
 * failure or by quitting, never leaves an unchecked file under the .dmg name.
 */
async function downloadUpdate(): Promise<void> {
  const release = offered
  if (release === undefined || status.state === 'downloading') return
  const { version, dmg } = release
  setStatus({ state: 'downloading', version, percent: 0 })
  let partPath: string | undefined
  let path: string
  let message = 'Couldn’t download the update. Try again later.'
  try {
    const sums = await (await request(release.checksumsUrl, 'text/plain', AbortSignal.timeout(15_000))).text()
    const expected = checksumFor(sums, dmg.name)
    if (expected === undefined) throw new Error(`SHA256SUMS.txt does not list ${dmg.name}.`)
    const downloads = app.getPath('downloads')
    const part = await claimName(downloads, `${dmg.name}.part`, candidate => open(candidate, 'wx'))
    partPath = part.path
    partInProgress = partPath
    const actual = await save(dmg, part.result, percent => setStatus({ state: 'downloading', version, percent }))
    if (actual !== expected) {
      message = 'The download didn’t match its checksum, so it was deleted. Try again.'
      throw new Error(`${dmg.name} has SHA-256 ${actual}; SHA256SUMS.txt lists ${expected}.`)
    }
    // A hard link fails when the name is taken, so an existing file is never replaced.
    const verifiedPart = partPath
    path = (await claimName(downloads, dmg.name, candidate => link(verifiedPart, candidate))).path
    await rm(verifiedPart, { force: true })
    partInProgress = undefined
  } catch (error) {
    console.warn('Update download failed:', error)
    if (partPath !== undefined) await rm(partPath, { force: true }).catch(() => undefined)
    partInProgress = undefined
    setStatus({ state: 'failed', version, message })
    return
  }
  // Mounts the dmg and shows its Finder window. A verified dmg stays in Downloads even if that fails.
  const openError = await shell.openPath(path)
  setStatus(openError === ''
    ? { state: 'downloaded', version }
    : { state: 'failed', version, message: `Open ${basename(path)} in Downloads to install it.` })
}

/** Answers the renderer's three update requests; status changes go to every window through `send`. */
export function registerUpdates(send: (channel: string, payload?: unknown) => void): void {
  broadcast = send
  app.on('will-quit', () => {
    if (partInProgress !== undefined) rmSync(partInProgress, { force: true })
  })
  handle('update:check', () => checkForUpdates())
  handle('update:download', () => downloadUpdate())
  handle('update:open-notes', async () => {
    // selectRelease only accepts a page under github.com/Zedyas/zen-forge/releases/.
    if (offered !== undefined) await shell.openExternal(offered.pageUrl)
  })
}

/**
 * Checks a few seconds after launch when automatic checks are on and the last check is a day old.
 * A failure stays silent and the next launch tries again. Development builds only check from the menu.
 */
export function scheduleUpdateCheck(): void {
  if (!app.isPackaged) return
  setTimeout(() => {
    const { checkForUpdates: enabled, lastUpdateCheck } = readSettings()
    // Either direction, so a clock that was set back does not stop checks.
    if (enabled && Math.abs(Date.now() - lastUpdateCheck) >= day) void checkForUpdates()
  }, 5_000)
}
