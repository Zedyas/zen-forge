import { toast } from 'sonner'
import { findCommand, type CommandId } from '@shared/commands'
import { errorMessage } from './feedback'
import { changeIndent, editTable, insertTable, setAlign, toggleList, toggleRunStyle, type TableEdit } from './format-actions'
import { newShape, newTextBox, type AlignEdge, type BasicShape, type HorizontalAlign } from './model'
import { exportPresentationPdf, printPresentation } from './print'
import { activeSlidesId, addElement, alignSelection, arrangeSelection, insertPicture, runSlideCommand } from './slides-actions'
import { readySlides, selectedElement } from './slides-store'

const shapeCommands: Partial<Record<CommandId, BasicShape>> = {
  'slide.insertRectangle': 'rect',
  'slide.insertRoundedRectangle': 'roundRect',
  'slide.insertEllipse': 'ellipse',
  'slide.insertTriangle': 'triangle',
  'slide.insertArrow': 'rightArrow',
  'slide.insertLine': 'line',
  'slide.insertLineArrow': 'arrow',
}

const alignCommands: Partial<Record<CommandId, AlignEdge>> = {
  'slide.alignLeft': 'left',
  'slide.alignCenter': 'center',
  'slide.alignRight': 'right',
  'slide.alignTop': 'top',
  'slide.alignMiddle': 'middle',
  'slide.alignBottom': 'bottom',
}

const tableCommands: Partial<Record<CommandId, TableEdit>> = {
  'slide.rowAbove': 'rowAbove',
  'slide.rowBelow': 'rowBelow',
  'slide.columnLeft': 'columnLeft',
  'slide.columnRight': 'columnRight',
  'slide.deleteRow': 'deleteRow',
  'slide.deleteColumn': 'deleteColumn',
}

const textAlignCommands: Partial<Record<CommandId, HorizontalAlign>> = {
  'text.alignLeft': 'left',
  'text.alignCenter': 'center',
  'text.alignRight': 'right',
}

/** Editor commands from the menu, keyboard and toolbar, applied to the active presentation. */
export async function runSlidesCommand(command: CommandId): Promise<void> {
  const id = activeSlidesId()
  const document = id === undefined ? undefined : readySlides(id)
  if (id === undefined || document === undefined) return
  try {
    const shape = shapeCommands[command]
    const edge = alignCommands[command]
    const tableEdit = tableCommands[command]
    const align = textAlignCommands[command]
    if (shape !== undefined) return addElement(id, newShape(shape, document.present))
    if (edge !== undefined) return alignSelection(id, edge)
    if (align !== undefined) return setAlign(id, align)
    if (tableEdit !== undefined) {
      if (selectedElement(document)?.kind === 'table') editTable(id, tableEdit)
      else toast.info('Select a table first')
      return
    }
    switch (command) {
      case 'format.bold':
        return toggleRunStyle(id, 'bold')
      case 'format.italic':
        return toggleRunStyle(id, 'italic')
      case 'format.underline':
        return toggleRunStyle(id, 'underline')
      case 'text.bullets':
        return toggleList(id, 'bullet')
      case 'text.numbering':
        return toggleList(id, 'number')
      case 'text.indent':
        return changeIndent(id, 1)
      case 'text.outdent':
        return changeIndent(id, -1)
      case 'slide.insertTextBox':
        return addElement(id, newTextBox(document.present), true)
      case 'slide.insertPicture':
        return await insertPicture(id)
      case 'slide.insertTable':
        return insertTable(id, 3, 3)
      case 'slide.bringToFront':
        return arrangeSelection(id, 'front')
      case 'slide.sendToBack':
        return arrangeSelection(id, 'back')
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
