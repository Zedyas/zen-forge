import { toast } from 'sonner'
import { findCommand, type CommandId } from '@shared/commands'
import { errorMessage } from './feedback'
import { toggleRunStyle } from './format-actions'
import { exportPresentationPdf, printPresentation } from './print'
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
      case 'file.print':
        return await printPresentation(id)
      case 'file.exportPdf':
        return await exportPresentationPdf(id)
      default:
        return await runSlideCommand(id, command)
    }
  } catch (error) {
    toast.error(`${findCommand(command).label.replace('…', '')} didn't work`, { description: errorMessage(error, 'The presentation was not changed.') })
  }
}
