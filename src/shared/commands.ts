import type { ApplicationId, EditorApplicationId } from './applications'

export const commandIds = [
  'file.new',
  'file.newDocument',
  'file.newPresentation',
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
  'format.tableOfContents',
  // Ledger
  'format.strikethrough',
  'format.wrapText',
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
  'edit.duplicate',
  'edit.delete',
  'text.bullets',
  'text.numbering',
  'text.indent',
  'text.outdent',
  'text.alignLeft',
  'text.alignCenter',
  'text.alignRight',
  'slide.new',
  'slide.duplicate',
  'slide.delete',
  'slide.moveUp',
  'slide.moveDown',
  'slide.insertTextBox',
  'slide.insertPicture',
  'slide.insertTable',
  'slide.insertRectangle',
  'slide.insertRoundedRectangle',
  'slide.insertEllipse',
  'slide.insertTriangle',
  'slide.insertArrow',
  'slide.insertLine',
  'slide.insertLineArrow',
  'slide.bringToFront',
  'slide.sendToBack',
  'slide.alignLeft',
  'slide.alignCenter',
  'slide.alignRight',
  'slide.alignTop',
  'slide.alignMiddle',
  'slide.alignBottom',
  'slide.rowAbove',
  'slide.rowBelow',
  'slide.columnLeft',
  'slide.columnRight',
  'slide.deleteRow',
  'slide.deleteColumn',
  'slide.play',
  'tab.previous',
  'tab.next',
  'tab.moveToNewWindow',
  'shell.home',
  'palette.open',
  'app.checkForUpdates',
] as const

export type CommandId = (typeof commandIds)[number]

/** `app` is the menu named after the app, which holds About and Quit. */
export type MenuId = 'app' | 'file' | 'edit' | 'format' | 'view' | 'page' | 'slide' | 'window' | 'none'

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
  /** Commands of one menu that share a group sit together in a submenu of that name. */
  readonly group?: string
}

export const commandDefinitions: readonly CommandDefinition[] = [
  { id: 'file.new', label: 'New Spreadsheet', menu: 'file', accelerator: 'CmdOrCtrl+N', scope: 'all', target: 'renderer' },
  { id: 'file.newDocument', label: 'New Document', menu: 'file', scope: 'all', target: 'renderer' },
  { id: 'file.newPresentation', label: 'New Presentation', menu: 'file', scope: 'all', target: 'renderer' },
  { id: 'window.new', label: 'New Window', menu: 'file', accelerator: 'Shift+CmdOrCtrl+N', scope: 'all', target: 'main' },
  { id: 'file.open', label: 'Open…', menu: 'file', accelerator: 'CmdOrCtrl+O', scope: 'all', target: 'renderer' },
  { id: 'file.save', label: 'Save', menu: 'file', accelerator: 'CmdOrCtrl+S', scope: 'editors', target: 'renderer', separatorBefore: true },
  { id: 'file.saveAs', label: 'Save As…', menu: 'file', accelerator: 'Shift+CmdOrCtrl+S', scope: 'editors', target: 'renderer' },
  { id: 'file.exportCsv', label: 'Export Sheet as CSV…', menu: 'file', scope: ['sheets'], target: 'renderer' },
  { id: 'file.exportPdf', label: 'Export as PDF…', menu: 'file', scope: ['sheets', 'docs', 'slides'], target: 'renderer' },
  { id: 'file.print', label: 'Print…', menu: 'file', accelerator: 'CmdOrCtrl+P', scope: ['sheets', 'pdf', 'docs', 'slides'], target: 'renderer', separatorBefore: true },
  { id: 'file.closeTab', label: 'Close Tab', menu: 'file', accelerator: 'CmdOrCtrl+W', scope: 'all', target: 'renderer', separatorBefore: true },
  { id: 'edit.undo', label: 'Undo', menu: 'edit', accelerator: 'CmdOrCtrl+Z', scope: 'editors', target: 'renderer' },
  { id: 'edit.redo', label: 'Redo', menu: 'edit', accelerator: 'Shift+CmdOrCtrl+Z', scope: 'editors', target: 'renderer' },
  { id: 'edit.find', label: 'Find…', menu: 'edit', accelerator: 'CmdOrCtrl+F', scope: ['sheets', 'pdf', 'docs', 'slides'], target: 'renderer', separatorBefore: true },
  { id: 'edit.duplicate', label: 'Duplicate', menu: 'edit', accelerator: 'CmdOrCtrl+D', scope: ['slides'], target: 'renderer' },
  { id: 'edit.delete', label: 'Delete', menu: 'edit', scope: ['slides'], target: 'renderer' },
  { id: 'format.bold', label: 'Bold', menu: 'format', accelerator: 'CmdOrCtrl+B', scope: ['sheets', 'docs', 'slides'], target: 'renderer' },
  { id: 'format.italic', label: 'Italic', menu: 'format', accelerator: 'CmdOrCtrl+I', scope: ['sheets', 'docs', 'slides'], target: 'renderer' },
  { id: 'format.underline', label: 'Underline', menu: 'format', accelerator: 'CmdOrCtrl+U', scope: ['sheets', 'docs', 'slides'], target: 'renderer' },
  { id: 'format.clear', label: 'Clear Formatting', menu: 'format', scope: ['sheets', 'docs'], target: 'renderer', separatorBefore: true },
  { id: 'format.tableOfContents', label: 'Insert Table of Contents', menu: 'format', scope: ['docs'], target: 'renderer', separatorBefore: true },
  { id: 'text.bullets', label: 'Bulleted List', menu: 'format', scope: ['slides'], target: 'renderer', separatorBefore: true },
  { id: 'text.numbering', label: 'Numbered List', menu: 'format', scope: ['slides'], target: 'renderer' },
  { id: 'text.indent', label: 'Increase Indent', menu: 'format', scope: ['slides'], target: 'renderer' },
  { id: 'text.outdent', label: 'Decrease Indent', menu: 'format', scope: ['slides'], target: 'renderer' },
  { id: 'text.alignLeft', label: 'Align Text Left', menu: 'format', scope: ['slides'], target: 'renderer', separatorBefore: true },
  { id: 'text.alignCenter', label: 'Align Text Centre', menu: 'format', scope: ['slides'], target: 'renderer' },
  { id: 'text.alignRight', label: 'Align Text Right', menu: 'format', scope: ['slides'], target: 'renderer' },
  // Ledger
  { id: 'format.strikethrough', label: 'Strikethrough', menu: 'format', accelerator: 'Shift+CmdOrCtrl+X', scope: ['sheets'], target: 'renderer' },
  { id: 'format.wrapText', label: 'Wrap Text', menu: 'format', scope: ['sheets'], target: 'renderer', separatorBefore: true },
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
  { id: 'slide.new', label: 'New Slide', menu: 'slide', accelerator: 'Alt+CmdOrCtrl+N', scope: ['slides'], target: 'renderer' },
  { id: 'slide.duplicate', label: 'Duplicate Slide', menu: 'slide', scope: ['slides'], target: 'renderer' },
  { id: 'slide.delete', label: 'Delete Slide', menu: 'slide', scope: ['slides'], target: 'renderer' },
  { id: 'slide.moveUp', label: 'Move Slide Up', menu: 'slide', scope: ['slides'], target: 'renderer' },
  { id: 'slide.moveDown', label: 'Move Slide Down', menu: 'slide', scope: ['slides'], target: 'renderer' },
  { id: 'slide.insertTextBox', label: 'Insert Text Box', menu: 'slide', scope: ['slides'], target: 'renderer', separatorBefore: true },
  { id: 'slide.insertPicture', label: 'Insert Picture…', menu: 'slide', scope: ['slides'], target: 'renderer' },
  { id: 'slide.insertTable', label: 'Insert Table', menu: 'slide', scope: ['slides'], target: 'renderer' },
  { id: 'slide.insertRectangle', label: 'Insert Rectangle', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Insert Shape' },
  { id: 'slide.insertRoundedRectangle', label: 'Insert Rounded Rectangle', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Insert Shape' },
  { id: 'slide.insertEllipse', label: 'Insert Ellipse', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Insert Shape' },
  { id: 'slide.insertTriangle', label: 'Insert Triangle', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Insert Shape' },
  { id: 'slide.insertArrow', label: 'Insert Arrow', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Insert Shape' },
  { id: 'slide.insertLine', label: 'Insert Line', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Insert Shape' },
  { id: 'slide.insertLineArrow', label: 'Insert Line with Arrowhead', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Insert Shape' },
  { id: 'slide.bringToFront', label: 'Bring to Front', menu: 'slide', scope: ['slides'], target: 'renderer', separatorBefore: true },
  { id: 'slide.sendToBack', label: 'Send to Back', menu: 'slide', scope: ['slides'], target: 'renderer' },
  { id: 'slide.alignLeft', label: 'Align Left Edges', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Align Objects' },
  { id: 'slide.alignCenter', label: 'Align Centres', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Align Objects' },
  { id: 'slide.alignRight', label: 'Align Right Edges', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Align Objects' },
  { id: 'slide.alignTop', label: 'Align Top Edges', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Align Objects' },
  { id: 'slide.alignMiddle', label: 'Align Middles', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Align Objects' },
  { id: 'slide.alignBottom', label: 'Align Bottom Edges', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Align Objects' },
  { id: 'slide.rowAbove', label: 'Insert Row Above', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Table', separatorBefore: true },
  { id: 'slide.rowBelow', label: 'Insert Row Below', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Table' },
  { id: 'slide.columnLeft', label: 'Insert Column Left', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Table' },
  { id: 'slide.columnRight', label: 'Insert Column Right', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Table' },
  { id: 'slide.deleteRow', label: 'Delete Table Row', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Table' },
  { id: 'slide.deleteColumn', label: 'Delete Table Column', menu: 'slide', scope: ['slides'], target: 'renderer', group: 'Table' },
  { id: 'slide.play', label: 'Play Slideshow', menu: 'slide', accelerator: 'Alt+CmdOrCtrl+P', scope: ['slides'], target: 'renderer', separatorBefore: true },
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
