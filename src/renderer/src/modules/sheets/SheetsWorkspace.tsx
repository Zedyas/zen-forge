import { useEffect } from 'react'
import type { CommandId } from '@shared/commands'
import { executeCommand } from '../../app/commands'
import type { OpenDocument } from '../../app/documents-store'
import type { EditorHandler } from '../../app/editors'
import { recordPreview } from '../../services/index/document-index'
import { WindowEmpty } from '../../ui/WindowEmpty'
import { renderPreview } from './preview'
import { activeSheet, toggleStyle } from './sheet-commands'
import { SheetEditor } from './SheetEditor'
import { exportSheetCsv, releaseWorkbook, saveWorkbook, useWorkbookEntry } from './workbook-documents'

/** The native menu owns ⌘Z, so text fields (formula bar, cell editor, find) get their own undo back here. */
function isEditingText(): boolean {
  const focused = document.activeElement
  return focused instanceof HTMLInputElement || focused instanceof HTMLTextAreaElement
}

function runSheetCommand(command: CommandId): void | Promise<void> {
  const controller = activeSheet()
  if (controller === undefined) return
  switch (command) {
    case 'edit.undo':
      if (isEditingText()) document.execCommand('undo')
      else controller.workbook.undo()
      return
    case 'edit.redo':
      if (isEditingText()) document.execCommand('redo')
      else controller.workbook.redo()
      return
    case 'edit.find':
      return controller.openFind()
    case 'format.bold':
      return toggleStyle(controller, 'bold')
    case 'format.italic':
      return toggleStyle(controller, 'italic')
    case 'format.underline':
      return toggleStyle(controller, 'underline')
    case 'format.clear':
      return controller.workbook.clearFormatting(controller.sheetId, controller.range)
    case 'file.exportCsv':
      return exportSheetCsv(controller.documentId, controller.sheetId, controller.sheetName)
  }
}

export const sheetsEditor: EditorHandler = { run: runSheetCommand, save: saveWorkbook, release: releaseWorkbook }

function LoadedDocument({ document }: { readonly document: OpenDocument }) {
  const entry = useWorkbookEntry(document)

  useEffect(() => {
    if (entry.status !== 'ready' || document.path === undefined) return
    const first = entry.workbook.sheets()[0]
    if (first !== undefined) void recordPreview(document.path, renderPreview(entry.workbook, first.id))
  }, [document.path, entry])

  if (entry.status === 'loading') return <section className="window-empty"><p>Opening {document.name}…</p></section>
  if (entry.status === 'error') {
    return (
      <WindowEmpty application="sheets" title={`${document.name} could not be opened`} description={entry.message}>
        <button type="button" className="button" onClick={() => void executeCommand('file.open')}>Open another file…</button>
      </WindowEmpty>
    )
  }
  return <SheetEditor key={document.id} document={document} workbook={entry.workbook} />
}

/** Ledger: the spreadsheet tab's content below the title bar. */
export function SheetsWorkspace({ document }: { readonly document: OpenDocument }) {
  return <LoadedDocument key={document.id} document={document} />
}
