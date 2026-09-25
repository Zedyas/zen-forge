import { useEffect } from 'react'
import type { OpenDocument } from '../../app/documents-store'
import type { EditorHandler } from '../../app/editors'
import { useViewStore } from '../../app/view-store'
import { FidelitySurface } from '../../ui/FidelitySurface'
import { TitleEssentials } from '../../ui/TitleSlot'
import { WindowEmpty } from '../../ui/WindowEmpty'
import { ConfirmHost } from '../../ui/ConfirmDialog'
import { closeFind, useFindStore } from './find-store'
import { FindBar } from './FindBar'
import { removeMarkup, findMarkup, updateMarkup } from './markup-actions'
import { translateMarkup } from './model'
import { PageRail } from './PageRail'
import { isTextEntry, loadPdfDocument, releasePdfDocument, runPdfCommand, savePdfDocument } from './pdf-actions'
import { PdfDesk } from './PdfDesk'
import { PdfInspector } from './PdfInspector'
import { readyPdf, unsavedChanges, usePdfStore, type ReadyPdf } from './pdf-store'
import { PdfEssentials, PdfToolbar } from './PdfToolbar'
import { SignatureDialog } from './SignatureDialog'
import { setTool, toolIds, toolKeys, useToolStore } from './tool-store'
import './pdf.css'

/** Single-letter tool keys, Delete, arrow nudges and Escape (which also closes find); ignored while typing. */
function useToolKeys(documentId: string | undefined): void {
  useEffect(() => {
    if (documentId === undefined) return
    const handleKeyDown = (event: KeyboardEvent): void => {
      if (event.metaKey || event.ctrlKey || event.altKey || isTextEntry(event.target)) return
      const key = event.key.toLowerCase()
      const { selection } = useToolStore.getState()
      if (key === 'escape') {
        setTool('select')
        useToolStore.setState({ selection: undefined })
        closeFind()
        return
      }
      if ((key === 'delete' || key === 'backspace') && selection !== undefined) {
        event.preventDefault()
        removeMarkup(documentId, selection.pageKey, selection.markupId)
        return
      }
      const nudge = { arrowleft: [-1, 0], arrowright: [1, 0], arrowup: [0, -1], arrowdown: [0, 1] }[key]
      if (nudge !== undefined && selection !== undefined) {
        const document = readyPdf(documentId)
        const markup = document === undefined ? undefined : findMarkup(document.present, selection.pageKey, selection.markupId)
        if (markup === undefined) return
        event.preventDefault()
        const step = event.shiftKey ? 10 : 1
        updateMarkup(documentId, selection.pageKey, selection.markupId, translateMarkup(markup, (nudge[0] ?? 0) * step, (nudge[1] ?? 0) * step))
        return
      }
      const tool = toolIds.find(candidate => toolKeys[candidate].toLowerCase() === key)
      if (tool !== undefined) setTool(tool)
    }
    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [documentId])
}

function StatusBar({ documentId, document }: { readonly documentId: string; readonly document: ReadyPdf }) {
  const changes = unsavedChanges(document)
  return (
    <footer className="statusbar">
      <span className="statusbar-item">Page <b>{document.currentPage + 1}</b> of {document.present.pages.length}</span>
      <span className="statusbar-item">{Math.round(document.zoom * 100)}%</span>
      {changes > 0 && <span className="statusbar-item">{changes} {changes === 1 ? 'change' : 'changes'} not saved</span>}
      <span className="statusbar-grow" />
      <FidelitySurface documentId={documentId} />
    </footer>
  )
}

export const pdfEditor: EditorHandler = { run: runPdfCommand, save: savePdfDocument, release: releasePdfDocument }

/** Hanko: the PDF tab's content below the title bar. */
export function PdfWorkspace({ document }: { readonly document: OpenDocument }) {
  const documentId = document.id
  const state = usePdfStore(store => store.documents[documentId])
  const toolbar = useViewStore(view => view.toolbar)
  const inspector = useViewStore(view => view.inspector)
  const findOpen = useFindStore(find => find.open)
  useToolKeys(state?.status === 'ready' ? documentId : undefined)

  useEffect(() => {
    if (document.path === undefined) return
    if (usePdfStore.getState().documents[document.id] !== undefined) return
    void loadPdfDocument(document.id, document.path, `${document.name}.${document.extension}`)
  }, [document])

  // Tool and find state are shared by every PDF tab; a different document starts with nothing selected and find closed.
  // Find also closes when the document does, which stops its search and drops its matches.
  useEffect(() => {
    useToolStore.setState({ selection: undefined, textDraft: undefined })
    closeFind()
    return closeFind
  }, [documentId])

  if (state === undefined || state.status === 'loading') {
    return <section className="window-empty"><p>Opening {document.name}…</p></section>
  }

  if (state.status === 'error') {
    return <WindowEmpty application="pdf" title={`${document.name} could not be opened`} description={state.message} />
  }

  return (
    <>
      <ConfirmHost />
      <SignatureDialog />
      <TitleEssentials><PdfEssentials /></TitleEssentials>
      {toolbar && <PdfToolbar documentId={documentId} document={state} />}
      {findOpen && <FindBar documentId={documentId} document={state} />}
      <div className="window-body">
        <PageRail documentId={documentId} document={state} />
        <div className="window-main">
          <PdfDesk key={documentId} documentId={documentId} document={state} />
          <StatusBar documentId={documentId} document={state} />
        </div>
        {inspector && <PdfInspector documentId={documentId} document={state} openDocument={document} />}
      </div>
    </>
  )
}
