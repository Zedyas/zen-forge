import { executeCommand } from '../../app/commands'
import type { OpenDocument } from '../../app/documents-store'
import type { EditorHandler } from '../../app/editors'
import { WindowEmpty } from '../../ui/WindowEmpty'
import { ConfirmHost } from '../../ui/ConfirmDialog'
import { runDocCommand } from './doc-commands'
import { releaseDoc, saveDoc, useDocEntry } from './doc-documents'
import { DocEditor } from './DocEditor'
import './docs.css'

export const docsEditor: EditorHandler = { run: runDocCommand, save: saveDoc, release: releaseDoc }

function LoadedDocument({ document }: { readonly document: OpenDocument }) {
  const entry = useDocEntry(document)
  if (entry.status === 'loading') return <section className="window-empty"><p>Opening {document.name}…</p></section>
  if (entry.status === 'error') {
    return (
      <WindowEmpty application="docs" title={`${document.name} could not be opened`} description={entry.message}>
        <button type="button" className="button" onClick={() => void executeCommand('file.open')}>Open another file…</button>
      </WindowEmpty>
    )
  }
  return <DocEditor key={document.id} document={document} editor={entry.editor} />
}

/** Sumi: the document tab's content below the title bar. */
export function DocsWorkspace({ document }: { readonly document: OpenDocument }) {
  return (
    <>
      <ConfirmHost />
      <LoadedDocument key={document.id} document={document} />
    </>
  )
}
