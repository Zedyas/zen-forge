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
  /**
   * Writes text still being typed (an open text box or field) into the document before it closes.
   * The tab is already marked unsaved from the first keystroke, so a window close or quit asks
   * first; this makes Save include the typing. Editors that keep typing in the document as it
   * happens omit it.
   */
  flush?(documentId: string): void
  /**
   * True while the document is presented full screen (a slideshow). Menu commands and shortcuts are
   * then ignored, so nothing changes the document mid-show; the show's own keys still work.
   */
  presenting?(documentId: string): boolean
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
