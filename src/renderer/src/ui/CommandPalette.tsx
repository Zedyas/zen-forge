import { Dialog } from '@base-ui/react/dialog'
import { Command } from 'cmdk'
import { Search } from 'lucide-react'
import { useEffect, useState } from 'react'
import { toast } from 'sonner'
import { suiteName } from '@shared/applications'
import { acceleratorLabel, commandApplies, commandDefinitions, type MenuId } from '@shared/commands'
import { recentMenuItems, type RecentMenuItem } from '@shared/recents'
import { executeCommand } from '../app/commands'
import { openRecent } from '../app/document-actions'
import { activeDocument, useDocumentsStore } from '../app/documents-store'
import { useShellUiStore } from '../app/shell-ui-store'
import { listRecents } from '../services/index/document-index'
import { icon } from './icons'

const groups: ReadonlyArray<readonly [MenuId, string]> = [
  ['file', 'File'],
  ['edit', 'Edit'],
  ['format', 'Format'],
  ['page', 'Page'],
  ['slide', 'Slide'],
  ['view', 'View'],
  ['window', 'Window'],
  ['app', suiteName],
]

/** Every command that applies to the active tab, searchable by name, with its shortcut. */
export function CommandPalette() {
  const open = useShellUiStore(state => state.paletteOpen)
  const setOpen = useShellUiStore(state => state.setPaletteOpen)
  const active = useDocumentsStore(state => activeDocument(state)?.kind ?? 'home')
  const [recents, setRecents] = useState<readonly RecentMenuItem[]>([])

  useEffect(() => {
    if (open) void listRecents().then(files => setRecents(recentMenuItems(files.map(file => file.path))))
  }, [open])

  return (
    <Dialog.Root open={open} onOpenChange={setOpen}>
      <Dialog.Portal>
        <Dialog.Backdrop className="palette-backdrop" />
        <Dialog.Popup className="palette-dialog">
          <Dialog.Title className="sr-only">Command palette</Dialog.Title>
          <Command className="command-palette" label="Command palette">
            <div className="command-input-row">
              <Search {...icon} />
              <Command.Input autoFocus placeholder="Type a command" />
              <kbd>esc</kbd>
            </div>
            <Command.List>
              <Command.Empty>No matching command.</Command.Empty>
              {groups.map(([menu, heading]) => {
                const commands = commandDefinitions.filter(command =>
                  command.menu === menu && commandApplies(command, active))
                if (commands.length === 0) return null
                return (
                  <Command.Group key={menu} heading={heading}>
                    {commands.map(command => (
                      <Command.Item
                        key={command.id}
                        value={`${heading} ${command.label}`}
                        onSelect={() => {
                          setOpen(false)
                          void executeCommand(command.id).catch(error => {
                            toast.error('Command failed', { description: error instanceof Error ? error.message : undefined })
                          })
                        }}
                      >
                        <span>{command.label.replace('…', '')}</span>
                        {command.accelerator !== undefined && <kbd>{acceleratorLabel(command.accelerator)}</kbd>}
                      </Command.Item>
                    ))}
                  </Command.Group>
                )
              })}
              {recents.length > 0 && (
                <Command.Group heading="Open Recent">
                  {recents.map(item => (
                    <Command.Item
                      key={item.path}
                      value={`Open Recent ${item.label}`}
                      onSelect={() => {
                        setOpen(false)
                        void openRecent(item.path).catch(error => {
                          toast.error('Could not open the file', { description: error instanceof Error ? error.message : undefined })
                        })
                      }}
                    >
                      <span>{item.label}</span>
                    </Command.Item>
                  ))}
                </Command.Group>
              )}
            </Command.List>
          </Command>
        </Dialog.Popup>
      </Dialog.Portal>
    </Dialog.Root>
  )
}
