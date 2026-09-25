import { useCallback, useEffect, useMemo, useState, type MouseEvent } from 'react'
import type { Editor } from '@tiptap/core'
import { EditorContent, useEditorState } from '@tiptap/react'
import type { OpenDocument } from '../../app/documents-store'
import { useViewStore } from '../../app/view-store'
import { FidelitySurface } from '../../ui/FidelitySurface'
import { openLink, setActiveDoc, type DocController } from './doc-commands'
import { page } from './doc-documents'
import { DocInspector } from './DocInspector'
import { DocEssentials, DocToolbar } from './DocToolbar'
import { clearFind } from './find'
import { FindBar } from './FindBar'
import { countText } from './text-stats'

function StatusBar({ document, editor }: { readonly document: OpenDocument; readonly editor: Editor }) {
  const { words, characters } = useEditorState({ editor, selector: ({ editor: current }) => countText(current.state.doc) })
  return (
    <footer className="statusbar">
      <span className="statusbar-item"><b>{words.toLocaleString()}</b> {words === 1 ? 'word' : 'words'}</span>
      <span className="statusbar-item"><b>{characters.toLocaleString()}</b> {characters === 1 ? 'character' : 'characters'}</span>
      <span className="statusbar-grow" />
      <FidelitySurface documentId={document.id} />
    </footer>
  )
}

/** ⌘-click follows a link, as in Word; a plain click places the caret. */
function followLink(event: MouseEvent): void {
  if (!event.metaKey || !(event.target instanceof Element)) return
  const href = event.target.closest('.sumi-prose a')?.getAttribute('href')
  if (href === null || href === undefined) return
  event.preventDefault()
  void openLink(href)
}

interface DocEditorProps {
  readonly document: OpenDocument
  readonly editor: Editor
}

/** One document on the desk: toolbar, find bar, the page, the status bar and the inspector. */
export function DocEditor({ document, editor }: DocEditorProps) {
  const toolbarShown = useViewStore(state => state.toolbar)
  const inspectorShown = useViewStore(state => state.inspector)
  // Counts ⌘F presses while the find bar is open, so each one focuses its field again.
  const [findRequest, setFindRequest] = useState<number | undefined>(undefined)
  const openFind = useCallback(() => setFindRequest(request => (request ?? 0) + 1), [])

  const controller = useMemo<DocController>(() => ({ documentId: document.id, editor, openFind }), [document.id, editor, openFind])
  useEffect(() => setActiveDoc(controller), [controller])
  // The caret is where it was when the tab was last shown; a new document is ready to type in.
  useEffect(() => {
    editor.commands.focus()
    // Find highlights belong to the open find bar; leaving the tab closes it.
    return () => clearFind(editor)
  }, [editor])

  const closeFind = (): void => {
    setFindRequest(undefined)
    clearFind(editor)
    editor.commands.focus()
  }

  return (
    <>
      <DocEssentials editor={editor} />
      {toolbarShown && <DocToolbar editor={editor} findOpen={findRequest !== undefined} onFind={openFind} />}
      <div className="window-body">
        <div className="window-main">
          {findRequest !== undefined && <FindBar editor={editor} focusRequest={findRequest} onClose={closeFind} />}
          <div className="sumi-desk" onClickCapture={followLink}>
            <EditorContent
              editor={editor}
              className="sumi-paper"
              style={{ width: `${page.width / 1440}in`, minHeight: `${page.height / 1440}in`, padding: `${page.margin / 1440}in` }}
            />
          </div>
          <StatusBar document={document} editor={editor} />
        </div>
        {inspectorShown && <DocInspector editor={editor} />}
      </div>
    </>
  )
}
