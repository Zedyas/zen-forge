import { app, BrowserWindow, dialog, Menu, nativeTheme, type IpcMainInvokeEvent, type MenuItemConstructorOptions } from 'electron'
import { basename, extname, join } from 'node:path'
import { applicationForExtension, findApplication, suiteName } from '../../src/shared/applications'
import { commandApplies, commandDefinitions, type CommandDefinition, type MenuId } from '../../src/shared/commands'
import type { Appearance, FileReference, ViewState, WindowSession, WindowState } from '../../src/shared/shell'
import { registerFileIpc } from './file-ipc'
import { NodeFileService } from './node-file-service'
import { registerPrintIpc } from './print-ipc'
import { applySecurity, appPageUrl, handle, registerAppScheme } from './security'
import { applyStoredAppearance, getAppearance, readSession, readSettings, setAppearance, setCheckForUpdates, writeSession } from './settings'
import { currentUpdateStatus, registerUpdates, scheduleUpdateCheck } from './updates'

/** Every window is the same kind: a row of tabs that starts on Home. */
interface WindowRecord {
  readonly window: BrowserWindow
  state: WindowState
  view: ViewState
  /** Last session's tabs for this window, handed to its renderer once. */
  restore: WindowSession | undefined
  /** Set once the renderer has resolved every unsaved document, so the next close proceeds. */
  closeApproved: boolean
}

const records: WindowRecord[] = []
const fileService = new NodeFileService()
/** Files handed over by Finder before the app finished launching. */
const pendingOpenPaths: string[] = []
let lastFocused: WindowRecord | undefined
let quitting = false

app.setName(suiteName)
registerAppScheme()

function recordFor(window: BrowserWindow | null | undefined): WindowRecord | undefined {
  return records.find(record => record.window === window)
}

/** Sends to a window's renderer, waiting for the page to load first; the preload queues until React listens. */
function send(record: WindowRecord, channel: string, payload?: unknown): void {
  const deliver = (): void => record.window.webContents.send(channel, payload)
  if (record.window.webContents.isLoading()) record.window.webContents.once('did-finish-load', deliver)
  else deliver()
}

function broadcast(channel: string, payload?: unknown): void {
  records.forEach(record => send(record, channel, payload))
}

/** The window the menu bar describes: the focused one, else the last focused, else the newest. */
function currentWindow(): WindowRecord | undefined {
  return recordFor(BrowserWindow.getFocusedWindow()) ?? lastFocused ?? records.at(-1)
}

/** The window a file or command from outside goes to; a new one when none is open. */
function targetWindow(): WindowRecord {
  const record = currentWindow() ?? createWindow()
  if (record.window.isMinimized()) record.window.restore()
  record.window.focus()
  return record
}

/** Remembers each open window's saved tabs for the next launch. Skipped while quitting so closing windows do not erase it. */
function saveSession(): void {
  if (quitting) return
  writeSession(records.map(record => record.state.session).filter(session => session.paths.length > 0))
}

function runCommand(command: CommandDefinition): void {
  if (command.id === 'window.new') {
    createWindow()
    return
  }
  send(targetWindow(), 'shell:command', command.id)
}

function menuItems(menu: MenuId, record: WindowRecord | undefined): MenuItemConstructorOptions[] {
  const active = record?.state.active ?? 'home'
  return commandDefinitions
    .filter(command => command.menu === menu && commandApplies(command, active))
    .flatMap(command => {
      const item: MenuItemConstructorOptions = {
        id: command.id,
        label: command.label,
        accelerator: command.accelerator,
        click: () => runCommand(command),
      }
      if (command.toggle !== undefined) {
        item.type = 'checkbox'
        item.checked = record?.view[command.toggle] ?? true
      }
      return command.separatorBefore === true ? [{ type: 'separator' as const }, item] : [item]
    })
}

function appearanceItems(): MenuItemConstructorOptions[] {
  const options: ReadonlyArray<readonly [Appearance, string]> = [['system', 'Use System Setting'], ['light', 'Light'], ['dark', 'Dark']]
  return options.map(([appearance, label]) => ({
    label,
    type: 'radio',
    checked: getAppearance() === appearance,
    click: () => {
      setAppearance(appearance)
      refreshMenu()
    },
  }))
}

/** The menu bar follows the focused window's active tab: Format appears for a spreadsheet, Page for a PDF. */
function buildMenu(record: WindowRecord | undefined): Menu {
  const edit = menuItems('edit', record)
  const find = edit.filter(item => item.id === 'edit.find')
  const format = menuItems('format', record)
  const page = menuItems('page', record)
  const view = menuItems('view', record)
  const toggles = view.filter(item => item.type === 'checkbox')
  const zoom = view.filter(item => item.type !== 'checkbox')
  const template: MenuItemConstructorOptions[] = [
    {
      label: suiteName,
      submenu: [
        { role: 'about', label: `About ${suiteName}` },
        ...menuItems('app', record),
        {
          label: 'Check for Updates Automatically',
          type: 'checkbox',
          checked: readSettings().checkForUpdates,
          click: item => setCheckForUpdates(item.checked),
        },
        { type: 'separator' },
        { role: 'services' },
        { type: 'separator' },
        { role: 'hide' },
        { role: 'hideOthers' },
        { role: 'unhide' },
        { type: 'separator' },
        { role: 'quit' },
      ],
    },
    {
      label: 'File',
      submenu: [...menuItems('file', record), { role: 'close', label: 'Close Window', accelerator: 'Shift+CmdOrCtrl+W' }],
    },
    {
      label: 'Edit',
      submenu: [
        ...edit.filter(item => item.id === 'edit.undo' || item.id === 'edit.redo'),
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { role: 'selectAll' },
        ...(find.length > 0 ? [{ type: 'separator' as const }, ...find] : []),
      ],
    },
    ...(format.length > 0 ? [{ label: 'Format', submenu: format }] : []),
    ...(page.length > 0 ? [{ label: 'Page', submenu: page }] : []),
    {
      label: 'View',
      submenu: [
        ...toggles,
        ...(toggles.length > 0 ? [{ type: 'separator' as const }] : []),
        { label: 'Appearance', submenu: appearanceItems() },
        ...zoom,
        { type: 'separator' },
        { role: 'togglefullscreen' },
        ...(app.isPackaged ? [] : [
          { type: 'separator' as const },
          { role: 'reload' as const, accelerator: '' },
          { role: 'toggleDevTools' as const, accelerator: 'Alt+CmdOrCtrl+J' },
        ]),
      ],
    },
    {
      label: 'Window',
      submenu: [
        { role: 'minimize' },
        { role: 'zoom' },
        { type: 'separator' },
        ...menuItems('window', record),
        { type: 'separator' },
        { role: 'front' },
      ],
    },
    { role: 'help', submenu: [] },
  ]
  return Menu.buildFromTemplate(template)
}

function refreshMenu(): void {
  Menu.setApplicationMenu(buildMenu(currentWindow()))
}

function loadRenderer(window: BrowserWindow): void {
  const developmentUrl = process.env['ELECTRON_RENDERER_URL']
  void window.loadURL(developmentUrl ?? appPageUrl)
}

function createWindow(restore?: WindowSession): WindowRecord {
  const window = new BrowserWindow({
    width: 1280,
    height: 840,
    minWidth: 860,
    minHeight: 540,
    show: false,
    title: suiteName,
    backgroundColor: nativeTheme.shouldUseDarkColors ? '#17181b' : '#f2f3f5',
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    // The buttons are 14pt circles and y is their top, so they centre on the title bar's centre line:
    // --titlebar-center (24px) in styles.css, minus 7.
    trafficLightPosition: process.platform === 'darwin' ? { x: 16, y: 17 } : undefined,
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      devTools: !app.isPackaged,
    },
  })

  const record: WindowRecord = {
    window,
    state: { title: suiteName, active: 'home', edited: false, hasUnsaved: false, session: restore ?? { paths: [] } },
    view: { toolbar: true, inspector: false },
    restore,
    closeApproved: false,
  }
  records.push(record)
  loadRenderer(window)
  const update = currentUpdateStatus()
  if (update.state !== 'none') send(record, 'update:status', update)

  window.once('ready-to-show', () => window.show())
  window.on('focus', () => {
    lastFocused = record
    refreshMenu()
  })
  window.on('close', event => {
    if (!record.state.hasUnsaved || record.closeApproved) return
    // The renderer walks its unsaved documents with Save / Don't Save / Cancel, then calls resolveClose.
    event.preventDefault()
    send(record, 'shell:close-requested')
  })
  window.on('closed', () => {
    records.splice(records.indexOf(record), 1)
    if (lastFocused === record) lastFocused = undefined
    saveSession()
    refreshMenu()
    // A quit paused for this window's unsaved documents resumes; the next window with changes asks again.
    if (quitting) app.quit()
  })

  return record
}

/** Opens a supported file as a tab in the target window. */
function openDocument(file: FileReference): void {
  const application = applicationForExtension(file.extension)
  if (application === undefined || !findApplication(application).available) return
  app.addRecentDocument(file.path)
  send(targetWindow(), 'shell:document-opened', file)
}

/** Finder has no window to show a toast in, so a file it hands over that cannot be read is reported in a message box. */
async function openPath(path: string): Promise<void> {
  if (applicationForExtension(extname(path).slice(1)) === undefined) return
  try {
    openDocument(await fileService.describeFile(path))
  } catch (error) {
    const settings = {
      type: 'warning' as const,
      message: `“${basename(path)}” could not be opened.`,
      detail: error instanceof Error ? error.message : undefined,
    }
    const owner = BrowserWindow.getFocusedWindow()
    await (owner === null ? dialog.showMessageBox(settings) : dialog.showMessageBox(owner, settings))
  }
}

function senderRecord(event: IpcMainInvokeEvent): WindowRecord | undefined {
  return recordFor(BrowserWindow.fromWebContents(event.sender))
}

handle('shell:new-window', () => {
  createWindow()
})
handle('shell:open-in-new-window', (_event, file: FileReference) => {
  send(createWindow(), 'shell:document-opened', file)
})
handle('shell:take-session', event => {
  const record = senderRecord(event)
  const restore = record?.restore
  if (record !== undefined) record.restore = undefined
  return restore
})
handle('shell:get-appearance', () => getAppearance())
handle('shell:set-appearance', (_event, appearance: Appearance) => {
  setAppearance(appearance)
  refreshMenu()
})
handle('shell:set-window-state', (event, state: WindowState) => {
  const record = senderRecord(event)
  if (record === undefined) return
  const activeChanged = record.state.active !== state.active
  record.state = state
  record.window.setTitle(state.title)
  if (process.platform === 'darwin') {
    record.window.setRepresentedFilename(state.path ?? '')
    record.window.setDocumentEdited(state.edited)
  }
  if (state.path !== undefined) app.addRecentDocument(state.path)
  saveSession()
  if (activeChanged && record === currentWindow()) refreshMenu()
})
handle('shell:set-view-state', (event, view: ViewState) => {
  const record = senderRecord(event)
  if (record === undefined) return
  record.view = view
  refreshMenu()
})
handle('shell:resolve-close', (event, approved: boolean) => {
  const record = senderRecord(event)
  if (record === undefined) return
  if (!approved) {
    quitting = false
    saveSession()
    return
  }
  record.closeApproved = true
  record.window.close()
})

// Finder double-click, Dock drops and "Open With" arrive here, possibly before the app is ready.
app.on('open-file', (event, path) => {
  event.preventDefault()
  if (app.isReady()) void openPath(path)
  else pendingOpenPaths.push(path)
})

app.on('before-quit', () => {
  quitting = true
})

app.whenReady().then(() => {
  // A packaged app takes its icon from the bundle; in development the Dock would show Electron's.
  if (!app.isPackaged) app.dock?.setIcon(join(__dirname, '../../build/icon.png'))
  applySecurity()
  applyStoredAppearance()
  registerFileIpc()
  registerUpdates(broadcast)
  registerPrintIpc()
  const sessions = readSession()
  if (sessions.length === 0) createWindow()
  else sessions.forEach(session => createWindow(session))
  refreshMenu()
  pendingOpenPaths.splice(0).forEach(path => void openPath(path))
  scheduleUpdateCheck()

  app.on('activate', () => {
    if (records.length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})
