import { useEffect } from 'react'
import { toast } from 'sonner'
import type { CommandId } from '@shared/commands'
import { platformClient } from '../services/platform/client'
import { executeCommand } from './commands'
import { useDocumentsStore } from './documents-store'
import { CommandPalette } from '../ui/CommandPalette'

function runCommand(command: CommandId): void {
  void executeCommand(command).catch(error => {
    toast.error('Command failed', {
      description: error instanceof Error ? error.message : 'The command could not be completed.',
    })
  })
}

/**
 * Handles the keys the native menu does not own: ⌘K for the palette, ⌘1–⌘9 to pick a tab (⌘9 is
 * the last, as in Safari) and ⌃Tab / ⌃⇧Tab to step through tabs.
 */
function handleKeyDown(event: KeyboardEvent): void {
  if (event.ctrlKey && event.key === 'Tab') {
    event.preventDefault()
    event.stopPropagation()
    useDocumentsStore.getState().selectRelative(event.shiftKey ? -1 : 1)
    return
  }
  if (!event.metaKey || event.shiftKey || event.altKey || event.ctrlKey) return
  if (event.key.toLowerCase() === 'k') {
    event.preventDefault()
    runCommand('palette.open')
  } else if (/^[1-9]$/.test(event.key)) {
    event.preventDefault()
    useDocumentsStore.getState().selectIndex(event.key === '9' ? Number.MAX_SAFE_INTEGER : Number(event.key) - 1)
  }
}

/** Connects native menu and keyboard command sources to the shared registry. */
export function CommandController() {
  useEffect(() => platformClient.onCommand(runCommand), [])

  useEffect(() => {
    // Capture phase, so the grid's own Tab handling does not swallow ⌃Tab first.
    window.addEventListener('keydown', handleKeyDown, true)
    return () => window.removeEventListener('keydown', handleKeyDown, true)
  }, [])

  return <CommandPalette />
}
