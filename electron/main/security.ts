import { app, ipcMain, net, protocol, session, type IpcMainInvokeEvent } from 'electron'
import { isAbsolute, join, relative, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

/*
 * Electron security checklist (electronjs.org/docs/latest/tutorial/security), applied in one place:
 * a private app:// scheme instead of file://, no navigation or new windows, no permissions beyond
 * writing to the clipboard, and IPC answered only for Zendo's own page.
 */

const scheme = 'app'
const host = 'zendo'
const rendererDir = join(__dirname, '../renderer')

/** The production page. Development loads Vite's dev server instead. */
export const appPageUrl = `${scheme}://${host}/index.html`

function developmentOrigin(): string | undefined {
  const url = process.env['ELECTRON_RENDERER_URL']
  return url === undefined ? undefined : new URL(url).origin
}

function isAppPage(url: string): boolean {
  const development = developmentOrigin()
  return url.startsWith(`${scheme}://${host}/`) || (development !== undefined && url.startsWith(`${development}/`))
}

/** Must run before the app is ready. `standard` and `secure` give the scheme an origin, storage and module scripts. */
export function registerAppScheme(): void {
  protocol.registerSchemesAsPrivileged([{ scheme, privileges: { standard: true, secure: true, supportFetchAPI: true } }])
}

/** Serves the built renderer folder and nothing else; a path that escapes it is refused. */
function serveAppScheme(): void {
  protocol.handle(scheme, request => {
    const url = new URL(request.url)
    const file = resolve(rendererDir, `.${decodeURIComponent(url.pathname)}`)
    const inside = relative(rendererDir, file)
    if (url.host !== host || inside.startsWith('..') || isAbsolute(inside)) return new Response('Not found', { status: 404 })
    return net.fetch(pathToFileURL(file).toString())
  })
}

/** The only permission Zendo's page uses is writing to the clipboard (copying cells). Electron grants everything else by default. */
function denyPermissions(): void {
  const allowed = (permission: string): boolean => permission === 'clipboard-sanitized-write'
  session.defaultSession.setPermissionRequestHandler((_contents, permission, callback) => callback(allowed(permission)))
  session.defaultSession.setPermissionCheckHandler((_contents, permission) => allowed(permission))
}

/** Pages stay on Zendo: no navigating away (a dropped file, a link), no pop-up windows, no embedded webviews. */
function lockNavigation(): void {
  app.on('web-contents-created', (_event, contents) => {
    contents.on('will-navigate', (event, url) => {
      if (!isAppPage(url)) event.preventDefault()
    })
    contents.setWindowOpenHandler(() => ({ action: 'deny' }))
    contents.on('will-attach-webview', event => event.preventDefault())
  })
}

/** Call once the app is ready. */
export function applySecurity(): void {
  serveAppScheme()
  denyPermissions()
}

lockNavigation()

/** `ipcMain.handle` that answers only Zendo's own page, never a navigated-away page or a sub-frame. */
export function handle<Args extends unknown[], Result>(
  channel: string,
  listener: (event: IpcMainInvokeEvent, ...args: Args) => Result,
): void {
  ipcMain.handle(channel, (event, ...args: Args) => {
    if (event.senderFrame === null || event.senderFrame.parent !== null || !isAppPage(event.senderFrame.url)) {
      throw new Error(`Refused ${channel}: not sent by Zendo's own page.`)
    }
    return listener(event, ...args)
  })
}
