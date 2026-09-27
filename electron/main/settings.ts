import { app, nativeTheme } from 'electron'
import { cpSync, existsSync, readdirSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { basename, dirname, join } from 'node:path'
import { suiteName } from '../../src/shared/applications'
import type { Appearance, WindowSession } from '../../src/shared/shell'

interface Settings {
  readonly appearance: Appearance
  /** Check GitHub for a newer release a few seconds after launch, at most once a day. */
  readonly checkForUpdates: boolean
  /** When a check last reached GitHub, in milliseconds since 1970; 0 before the first. */
  readonly lastUpdateCheck: number
}

const defaults: Settings = { appearance: 'system', checkForUpdates: true, lastUpdateCheck: 0 }

/**
 * Zen Suzu was called Zendo before 0.2.1, and Electron names the data folder after the app. The first
 * launch under the new name copies the Zendo folder, so settings, open tabs and the page's storage
 * (recents, saved signatures) carry over. The Zendo folder is left as it was.
 * Electron creates the data folder, empty, before this runs, so an empty folder means a first launch.
 * The copy goes to a side folder first and then replaces the empty one, so a copy cut short is never
 * used. A --user-data-dir launch (tests) only copies from a Zendo folder beside it.
 */
export function copyZendoData(): void {
  const current = app.getPath('userData')
  const previous = join(dirname(current), 'Zendo')
  const used = existsSync(current) && readdirSync(current).length > 0
  if (basename(current) !== suiteName || used || !existsSync(previous)) return
  const staging = `${current}.copying`
  try {
    rmSync(staging, { recursive: true, force: true })
    cpSync(previous, staging, { recursive: true })
    // rename() replaces an empty folder.
    renameSync(staging, current)
  } catch {
    // Start with defaults instead. The Zendo folder is untouched; a partial copy stays in the side folder, unused.
  }
}

function settingsPath(): string {
  return join(app.getPath('userData'), 'settings.json')
}

function isAppearance(value: unknown): value is Appearance {
  return value === 'system' || value === 'light' || value === 'dark'
}

/** Each field is checked on its own; a missing or invalid one falls back to its default. */
export function readSettings(): Settings {
  try {
    const parsed: unknown = JSON.parse(readFileSync(settingsPath(), 'utf8'))
    if (typeof parsed !== 'object' || parsed === null) return defaults
    return {
      appearance: 'appearance' in parsed && isAppearance(parsed.appearance) ? parsed.appearance : defaults.appearance,
      checkForUpdates: 'checkForUpdates' in parsed && typeof parsed.checkForUpdates === 'boolean'
        ? parsed.checkForUpdates
        : defaults.checkForUpdates,
      lastUpdateCheck: 'lastUpdateCheck' in parsed && typeof parsed.lastUpdateCheck === 'number' && Number.isFinite(parsed.lastUpdateCheck)
        ? parsed.lastUpdateCheck
        : defaults.lastUpdateCheck,
    }
  } catch {
    // A missing or unreadable settings file means defaults; it is rewritten on the next change.
    return defaults
  }
}

function writeSettings(change: Partial<Settings>): void {
  writeFileSync(settingsPath(), JSON.stringify({ ...readSettings(), ...change }, null, 2))
}

/** Appearance drives `nativeTheme`, so CSS `prefers-color-scheme`, menus and window chrome all follow it. */
export function applyStoredAppearance(): void {
  nativeTheme.themeSource = readSettings().appearance
}

export function getAppearance(): Appearance {
  return nativeTheme.themeSource
}

export function setAppearance(appearance: Appearance): void {
  if (!isAppearance(appearance)) throw new Error('Unknown appearance.')
  nativeTheme.themeSource = appearance
  writeSettings({ appearance })
}

export function setCheckForUpdates(enabled: boolean): void {
  writeSettings({ checkForUpdates: enabled })
}

export function setLastUpdateCheck(time: number): void {
  try {
    writeSettings({ lastUpdateCheck: time })
  } catch {
    // Not persisting only means the next launch checks again.
  }
}

function sessionPath(): string {
  return join(app.getPath('userData'), 'session.json')
}

function isWindowSession(value: unknown): value is WindowSession {
  if (typeof value !== 'object' || value === null || !('paths' in value)) return false
  const { paths } = value
  const activePath = 'activePath' in value ? value.activePath : undefined
  return Array.isArray(paths) && paths.every(path => typeof path === 'string')
    && (activePath === undefined || typeof activePath === 'string')
}

/** The windows open at the last quit, each with its saved tabs. */
export function readSession(): readonly WindowSession[] {
  try {
    const parsed: unknown = JSON.parse(readFileSync(sessionPath(), 'utf8'))
    return Array.isArray(parsed) ? parsed.filter(isWindowSession) : []
  } catch {
    // No session file yet, or an unreadable one: start with one empty window.
    return []
  }
}

export function writeSession(sessions: readonly WindowSession[]): void {
  try {
    writeFileSync(sessionPath(), JSON.stringify(sessions))
  } catch {
    // Not persisting only costs reopening these tabs on the next launch.
  }
}
