import type { EditorApplicationId } from '@shared/applications'
import type { CommandId } from '@shared/commands'
import { docsEditor } from '../modules/docs/DocsWorkspace'
import { pdfEditor } from '../modules/pdf/PdfWorkspace'
import { sheetsEditor } from '../modules/sheets/SheetsWorkspace'
import { slidesEditor } from '../modules/slides/SlidesWorkspace'

/** What the shell needs from an editor module, for any of its documents, mounted or not. */
export interface EditorHandler {
  /** Runs an editor command (undo, format, page, zoom, print…) against the active document. */
  run(command: CommandId): void | Promise<void>
  /** Saves a document, asking for a path when it has none or `saveAs` is set. Resolves false when cancelled. */
  save(documentId: string, saveAs: boolean): Promise<boolean>
  /** Drops the module's in-memory state for a closed document. */
  release(documentId: string): void
}

/** The module that owns a document kind. Resolved at call time, so the import cycle with the modules is harmless. */
export function editorFor(kind: EditorApplicationId): EditorHandler | undefined {
  switch (kind) {
    case 'sheets':
      return sheetsEditor
    case 'pdf':
      return pdfEditor
    case 'docs':
      return docsEditor
    case 'slides':
      return slidesEditor
    default:
      return undefined
  }
}
