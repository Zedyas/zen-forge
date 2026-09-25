import { useEffect, useRef, useState, type KeyboardEvent } from 'react'
import type { Editor } from '@tiptap/core'
import { useEditorState } from '@tiptap/react'
import { CaseSensitive, ChevronDown, ChevronUp, X } from 'lucide-react'
import { toast } from 'sonner'
import { ToolButton } from '../../ui/Toolbar'
import { findState, moveFind, replaceAll, replaceCurrent, setFindQuery } from './find'

interface FindBarProps {
  readonly editor: Editor
  /** Changes on every ⌘F, which focuses the find field again. */
  readonly focusRequest: number
  onClose(): void
}

/** Selected text of one line starts the search, as in macOS apps. */
function selectedText(editor: Editor): string {
  const { from, to, empty } = editor.state.selection
  const text = empty ? '' : editor.state.doc.textBetween(from, to, '\n')
  return text.includes('\n') || text.length > 200 ? '' : text
}

/** Find and replace in the document; matches are highlighted on the page. */
export function FindBar({ editor, focusRequest, onClose }: FindBarProps) {
  const [query, setQuery] = useState(() => selectedText(editor))
  const [replacement, setReplacement] = useState('')
  const [caseSensitive, setCaseSensitive] = useState(false)
  const input = useRef<HTMLInputElement>(null)
  const { count, current } = useEditorState({
    editor,
    selector: ({ editor: state }) => {
      const find = findState(state.state)
      return { count: find?.matches.length ?? 0, current: find?.current ?? 0 }
    },
  })

  useEffect(() => {
    input.current?.focus()
    input.current?.select()
  }, [focusRequest])

  useEffect(() => setFindQuery(editor, query, caseSensitive), [caseSensitive, editor, query])

  const keyDown = (event: KeyboardEvent<HTMLInputElement>, onEnter: () => void): void => {
    if (event.key === 'Enter') {
      event.preventDefault()
      onEnter()
    }
    if (event.key === 'Escape') onClose()
  }

  return (
    <div className="find-bar" role="search">
      <input
        ref={input}
        className="text-field"
        value={query}
        placeholder="Find"
        aria-label="Find"
        onChange={event => setQuery(event.currentTarget.value)}
        onKeyDown={event => keyDown(event, () => moveFind(editor, event.shiftKey ? -1 : 1))}
      />
      <ToolButton icon={CaseSensitive} label="Match case" pressed={caseSensitive} onClick={() => setCaseSensitive(value => !value)} />
      <output className="find-count">{query === '' ? '' : count === 0 ? 'No matches' : `${current + 1} of ${count}`}</output>
      <ToolButton icon={ChevronUp} label="Previous match" disabled={count === 0} onClick={() => moveFind(editor, -1)} />
      <ToolButton icon={ChevronDown} label="Next match" disabled={count === 0} onClick={() => moveFind(editor, 1)} />
      <span className="tool-sep" />
      <input
        className="text-field"
        value={replacement}
        placeholder="Replace with"
        aria-label="Replace with"
        onChange={event => setReplacement(event.currentTarget.value)}
        onKeyDown={event => keyDown(event, () => replaceCurrent(editor, replacement))}
      />
      <button type="button" className="button" disabled={count === 0} onClick={() => replaceCurrent(editor, replacement)}>Replace</button>
      <button
        type="button"
        className="button"
        disabled={count === 0}
        onClick={() => {
          const replaced = replaceAll(editor, replacement)
          toast.success(`Replaced ${replaced} ${replaced === 1 ? 'match' : 'matches'}`)
        }}
      >
        Replace all
      </button>
      <span className="toolbar-grow" />
      <ToolButton icon={X} label="Close find" onClick={onClose} />
    </div>
  )
}
