import type { ApplicationId, EditorApplicationId } from './applications'

export const commandIds = [
  'file.new',
  'file.newDocument',
  'window.new',
  'file.open',
  'file.save',
  'file.saveAs',
  'file.exportCsv',
  'file.exportPdf',
  'file.print',
  'file.closeTab',
  'edit.undo',
  'edit.redo',
  'edit.find',
  'format.bold',
  'format.italic',
  'format.underline',
  'format.clear',
  'view.toolbar',
  'view.inspector',
  'view.zoomIn',
  'view.zoomOut',
  'view.actualSize',
  'page.rotateLeft',
  'page.rotateRight',
  'page.delete',
  'page.insert',
  'page.extract',
  'tab.previous',
  'tab.next',
  'tab.moveToNewWindow',
  'shell.home',
  'palette.open',
  'app.checkForUpdates',
] as const

export type CommandId = (typeof commandIds)[number]

/** `app` is the menu named after the app, which holds About and Quit. */
export type MenuId = 'app' | 'file' | 'edit' | 'format' | 'view' | 'page' | 'window' | 'none'

export interface CommandDefinition {
  readonly id: CommandId
  readonly label: string
  readonly menu: MenuId
  readonly accelerator?: string
  /**
   * The active tabs the command applies to: `all` includes Home, `editors` means any document tab,
   * a list names the applications.
   */
  readonly scope: 'all' | 'editors' | readonly EditorApplicationId[]
  /** `main` commands act on windows; everything else is forwarded to the focused window's renderer. */
  readonly target: 'main' | 'renderer'
  /** Draws a check mark bound to the window's view state. */
  readonly toggle?: 'toolbar' | 'inspector'
  /** Starts a new group in its menu. */
  readonly separatorBefore?: boolean
}

export const commandDefinitions: readonly CommandDefinition[] = [
  { id: 'file.new', label: 'New Spreadsheet', menu: 'file', accelerator: 'CmdOrCtrl+N', scope: 'all', target: 'renderer' },
  { id: 'file.newDocument', label: 'New Document', menu: 'file', scope: 'all', target: 'renderer' },
  { id: 'window.new', label: 'New Window', menu: 'file', accelerator: 'Shift+CmdOrCtrl+N', scope: 'all', target: 'main' },
  { id: 'file.open', label: 'Open…', menu: 'file', accelerator: 'CmdOrCtrl+O', scope: 'all', target: 'renderer' },
  { id: 'file.save', label: 'Save', menu: 'file', accelerator: 'CmdOrCtrl+S', scope: 'editors', target: 'renderer', separatorBefore: true },
  { id: 'file.saveAs', label: 'Save As…', menu: 'file', accelerator: 'Shift+CmdOrCtrl+S', scope: 'editors', target: 'renderer' },
  { id: 'file.exportCsv', label: 'Export Sheet as CSV…', menu: 'file', scope: ['sheets'], target: 'renderer' },
  { id: 'file.exportPdf', label: 'Export as PDF…', menu: 'file', scope: ['sheets'], target: 'renderer' },
  { id: 'file.print', label: 'Print…', menu: 'file', accelerator: 'CmdOrCtrl+P', scope: ['sheets', 'pdf'], target: 'renderer', separatorBefore: true },
  { id: 'file.closeTab', label: 'Close Tab', menu: 'file', accelerator: 'CmdOrCtrl+W', scope: 'all', target: 'renderer', separatorBefore: true },
  { id: 'edit.undo', label: 'Undo', menu: 'edit', accelerator: 'CmdOrCtrl+Z', scope: 'editors', target: 'renderer' },
  { id: 'edit.redo', label: 'Redo', menu: 'edit', accelerator: 'Shift+CmdOrCtrl+Z', scope: 'editors', target: 'renderer' },
  { id: 'edit.find', label: 'Find…', menu: 'edit', accelerator: 'CmdOrCtrl+F', scope: ['sheets', 'pdf', 'docs'], target: 'renderer', separatorBefore: true },
  { id: 'format.bold', label: 'Bold', menu: 'format', accelerator: 'CmdOrCtrl+B', scope: ['sheets', 'docs'], target: 'renderer' },
  { id: 'format.italic', label: 'Italic', menu: 'format', accelerator: 'CmdOrCtrl+I', scope: ['sheets', 'docs'], target: 'renderer' },
  { id: 'format.underline', label: 'Underline', menu: 'format', accelerator: 'CmdOrCtrl+U', scope: ['sheets', 'docs'], target: 'renderer' },
  { id: 'format.clear', label: 'Clear Formatting', menu: 'format', scope: ['sheets', 'docs'], target: 'renderer', separatorBefore: true },
  { id: 'view.toolbar', label: 'Toolbar', menu: 'view', accelerator: 'Alt+CmdOrCtrl+T', scope: 'editors', target: 'renderer', toggle: 'toolbar' },
  { id: 'view.inspector', label: 'Inspector', menu: 'view', accelerator: 'Alt+CmdOrCtrl+I', scope: 'editors', target: 'renderer', toggle: 'inspector' },
  { id: 'view.zoomIn', label: 'Zoom In', menu: 'view', accelerator: 'CmdOrCtrl+=', scope: ['pdf', 'docs'], target: 'renderer', separatorBefore: true },
  { id: 'view.zoomOut', label: 'Zoom Out', menu: 'view', accelerator: 'CmdOrCtrl+-', scope: ['pdf', 'docs'], target: 'renderer' },
  { id: 'view.actualSize', label: 'Actual Size', menu: 'view', accelerator: 'CmdOrCtrl+0', scope: ['pdf', 'docs'], target: 'renderer' },
  { id: 'page.rotateLeft', label: 'Rotate Left', menu: 'page', accelerator: 'CmdOrCtrl+L', scope: ['pdf'], target: 'renderer' },
  { id: 'page.rotateRight', label: 'Rotate Right', menu: 'page', accelerator: 'CmdOrCtrl+R', scope: ['pdf'], target: 'renderer' },
  { id: 'page.delete', label: 'Delete Page', menu: 'page', accelerator: 'CmdOrCtrl+Backspace', scope: ['pdf'], target: 'renderer' },
  { id: 'page.insert', label: 'Insert Pages from File…', menu: 'page', scope: ['pdf'], target: 'renderer', separatorBefore: true },
  { id: 'page.extract', label: 'Extract Page…', menu: 'page', scope: ['pdf'], target: 'renderer' },
  { id: 'tab.previous', label: 'Show Previous Tab', menu: 'window', accelerator: 'Shift+CmdOrCtrl+[', scope: 'all', target: 'renderer' },
  { id: 'tab.next', label: 'Show Next Tab', menu: 'window', accelerator: 'Shift+CmdOrCtrl+]', scope: 'all', target: 'renderer' },
  { id: 'tab.moveToNewWindow', label: 'Move Tab to New Window', menu: 'window', scope: 'editors', target: 'renderer' },
  { id: 'shell.home', label: 'Home', menu: 'window', accelerator: 'Shift+CmdOrCtrl+H', scope: 'all', target: 'renderer', separatorBefore: true },
  { id: 'palette.open', label: 'Command Palette…', menu: 'none', accelerator: 'CmdOrCtrl+K', scope: 'all', target: 'renderer' },
  { id: 'app.checkForUpdates', label: 'Check for Updates…', menu: 'app', scope: 'all', target: 'renderer' },
]

export function findCommand(id: CommandId): CommandDefinition {
  const command = commandDefinitions.find(candidate => candidate.id === id)
  if (command === undefined) throw new Error(`Unknown command: ${id}`)
  return command
}

/** Whether a command can run while `active` is the window's active tab (`home` for the Home tab). */
export function commandApplies(command: CommandDefinition, active: ApplicationId): boolean {
  if (command.scope === 'all') return true
  if (active === 'home') return false
  return command.scope === 'editors' || command.scope.includes(active)
}

/** `Shift+CmdOrCtrl+S` → `⇧⌘S`, the form macOS menus and the palette show. */
export function acceleratorLabel(accelerator: string): string {
  const symbols: Record<string, string> = {
    CmdOrCtrl: '⌘', Shift: '⇧', Alt: '⌥', Backspace: '⌫', '=': '+',
  }
  return accelerator.split('+').map(part => symbols[part] ?? part).join('')
}
