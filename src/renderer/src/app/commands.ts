import { commandApplies, findCommand, type CommandId } from '@shared/commands'
import { platformClient } from '../services/platform/client'
import { chooseAndOpenFiles, closeDocument, moveToNewWindow, saveDocument } from './document-actions'
import { activeDocument, homeTabId, useDocumentsStore } from './documents-store'
import { editorFor } from './editors'
import { useShellUiStore } from './shell-ui-store'
import { useViewStore } from './view-store'

/** Runs a command from the native menu, the keyboard, the palette or a toolbar button. */
export async function executeCommand(id: CommandId): Promise<void> {
  const command = findCommand(id)
  const documents = useDocumentsStore.getState()
  const active = activeDocument(documents)
  if (!commandApplies(command, active?.kind ?? 'home')) throw new Error(`${command.label} is not available here.`)

  switch (id) {
    case 'palette.open':
      useShellUiStore.getState().setPaletteOpen(true)
      return
    case 'shell.home':
      documents.select(homeTabId)
      return
    case 'window.new':
      return platformClient.newWindow()
    case 'file.new':
      documents.openUntitled('sheets')
      return
    case 'file.open':
      return chooseAndOpenFiles()
    case 'file.closeTab':
      // Home cannot close, so ⌘W on it closes the window, as on a browser's last tab.
      if (active === undefined) window.close()
      else await closeDocument(active.id)
      return
    case 'tab.previous':
    case 'tab.next':
      documents.selectRelative(id === 'tab.next' ? 1 : -1)
      return
    case 'tab.moveToNewWindow':
      if (active !== undefined) await moveToNewWindow(active.id)
      return
    case 'view.toolbar':
      useViewStore.getState().toggle('toolbar')
      return
    case 'view.inspector':
      useViewStore.getState().toggle('inspector')
      return
    case 'file.save':
    case 'file.saveAs':
      if (active !== undefined) await saveDocument(active.id, id === 'file.saveAs')
      return
    default:
      if (active !== undefined) await editorFor(active.kind)?.run(id)
  }
}
