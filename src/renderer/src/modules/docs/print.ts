import type { Editor } from '@tiptap/core'
import { DOMSerializer } from '@tiptap/pm/model'
import { toast } from 'sonner'
import { openFiles } from '../../app/document-actions'
import { useDocumentsStore } from '../../app/documents-store'
import { fileService } from '../../services/file/IpcFileService'
import { openPrintDialog, printToPdf, type Printout } from '../../services/print/print'
import { documentPage, documentTheme, themeVariables } from './theme'

const millimetresPerTwip = 25.4 / 1440

/**
 * The document as it prints: the editor's content serialized by the schema, styled by the same
 * `.sumi-prose` rules and theme variables as the page on screen (docs.css), on the document's own
 * paper. Word margins can differ on each side, and page numbers sit in the bottom margin, so the
 * printout carries its own @page rule for both; the print service's rule sets the paper size.
 */
export function docPrintout(editor: Editor): Printout {
  const doc = editor.state.doc
  const page = documentPage(doc.attrs['page'])
  const theme = documentTheme(doc.attrs['theme'])
  const margin = (twips: number): string => `${(twips * millimetresPerTwip).toFixed(2)}mm`
  const pageNumbers = page.pageNumbers
    ? `@bottom-center { content: counter(page); font: 9pt "${theme.normal.fontFamily}", sans-serif; color: #555; }`
    : ''
  const rule = document.createElement('style')
  rule.textContent = `@page { margin: ${margin(page.marginTop)} ${margin(page.marginRight)} ${margin(page.marginBottom)} ${margin(page.marginLeft)}; ${pageNumbers} }`

  const prose = document.createElement('div')
  prose.className = 'sumi-prose sumi-print'
  for (const [name, value] of themeVariables(theme)) prose.style.setProperty(name, value)
  prose.append(DOMSerializer.fromSchema(editor.schema).serializeFragment(doc.content))

  const content = document.createElement('div')
  content.append(rule, prose)
  return { content, page: { width: page.width * millimetresPerTwip, height: page.height * millimetresPerTwip, margin: 0 } }
}

export async function printDoc(editor: Editor): Promise<void> {
  await openPrintDialog(() => docPrintout(editor))
}

/** Writes the document, as it prints, to a PDF chosen by the user, and offers to open it in Hanko. */
export async function exportDocPdf(documentId: string, editor: Editor): Promise<void> {
  const document = useDocumentsStore.getState().documents.find(candidate => candidate.id === documentId)
  if (document === undefined) return
  const path = await fileService.chooseSavePath({ defaultName: `${document.name}.pdf`, extensions: ['pdf'] })
  if (path === undefined) return
  const bytes = await printToPdf(() => docPrintout(editor))
  if (bytes === undefined) return
  await fileService.write(path, bytes)
  const file = await fileService.describe(path)
  toast.success('Exported as PDF', {
    description: file.name,
    action: { label: 'Open', onClick: () => void openFiles([file]) },
  })
}
