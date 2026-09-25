import { toast } from 'sonner'
import type { CommandId } from '@shared/commands'
import { toggleRunStyle } from './format-actions'
import { activeSlidesId, runSlideCommand } from './slides-actions'

/** Editor commands from the menu, keyboard and toolbar, applied to the active presentation. */
export async function runSlidesCommand(command: CommandId): Promise<void> {
  const id = activeSlidesId()
  if (id === undefined) return
  try {
    switch (command) {
      case 'format.bold':
        return toggleRunStyle(id, 'bold')
      case 'format.italic':
        return toggleRunStyle(id, 'italic')
      case 'format.underline':
        return toggleRunStyle(id, 'underline')
      default:
        return await runSlideCommand(id, command)
    }
  } catch (error) {
    toast.error('Could not complete that', { description: error instanceof Error && error.message !== '' ? error.message : 'The presentation could not be changed.' })
  }
}
