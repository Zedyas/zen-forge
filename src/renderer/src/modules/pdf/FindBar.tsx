import { ChevronDown, ChevronUp, X } from 'lucide-react'
import { useEffect, useRef, useState } from 'react'
import { toast } from 'sonner'
import { Tip } from '../../ui/Tip'
import { ToolButton } from '../../ui/Toolbar'
import { closeFind, findInDocument, redactMatches, stepFind, useFindStore } from './find-store'
import type { ReadyPdf } from './pdf-store'

interface FindBarProps {
  readonly documentId: string
  readonly document: ReadyPdf
}

function countLabel(query: string, found: number, current: number, searching: boolean): string {
  if (query.trim() === '') return ''
  if (found === 0) return searching ? 'Searching…' : 'No matches'
  return `${current + 1} of ${found}${searching ? '…' : ''}`
}

/** Find in the page text, case-insensitive, with a button that marks every match for redaction. */
export function FindBar({ documentId, document }: FindBarProps) {
  const query = useFindStore(state => state.query)
  const matches = useFindStore(state => state.matches)
  const current = useFindStore(state => state.current)
  const searching = useFindStore(state => state.searching)
  const focusRequest = useFindStore(state => state.focusRequest)
  const input = useRef<HTMLInputElement>(null)
  /** Set while "Redact all" asks MuPDF where the matches are, which loads it the first time. */
  const [redacting, setRedacting] = useState(false)
  // Reordering, rotating, adding or removing pages moves the matches; markups do not.
  const layout = document.present.pages.map(page => `${page.key}:${page.rotation}`).join(' ')

  useEffect(() => {
    input.current?.focus()
    input.current?.select()
  }, [focusRequest])

  useEffect(() => {
    void findInDocument(documentId, useFindStore.getState().query, false)
  }, [documentId, layout])

  const redactAll = async (): Promise<void> => {
    setRedacting(true)
    try {
      const marked = await redactMatches(documentId)
      if (marked === 0) toast.info('Every match is already marked for redaction')
      else if (marked !== undefined) toast.success(`Marked ${marked} ${marked === 1 ? 'match' : 'matches'} for redaction`, { description: 'They are removed from the file when you save.' })
    } catch (error) {
      toast.error('Could not mark the matches', { description: error instanceof Error ? error.message : 'Try Redact all again.' })
    } finally {
      setRedacting(false)
    }
  }

  return (
    <div className="find-bar" role="search">
      <input
        ref={input}
        className="text-field"
        value={query}
        placeholder="Find"
        aria-label="Find"
        onChange={event => void findInDocument(documentId, event.currentTarget.value, true)}
        onKeyDown={event => {
          if (event.key === 'Enter') stepFind(event.shiftKey ? -1 : 1)
          if (event.key === 'Escape') closeFind()
        }}
      />
      <output className="find-count">{countLabel(query, matches.length, current, searching)}</output>
      <ToolButton icon={ChevronUp} label="Previous match" disabled={matches.length === 0} onClick={() => stepFind(-1)} />
      <ToolButton icon={ChevronDown} label="Next match" disabled={matches.length === 0} onClick={() => stepFind(1)} />
      <span className="tool-sep" />
      <Tip label="Cover every match with a redaction box">
        <button type="button" className="button" disabled={matches.length === 0 || searching || redacting} onClick={() => void redactAll()}>
          Redact all
        </button>
      </Tip>
      <span className="toolbar-grow" />
      <ToolButton icon={X} label="Close find" onClick={closeFind} />
    </div>
  )
}

/** One page's matches, over the page image and under the markups; they never take the pointer. */
export function FindHighlights({ pageKey, zoom }: { readonly pageKey: string; readonly zoom: number }) {
  const matches = useFindStore(state => state.matches)
  const current = useFindStore(state => state.current)
  const hits = matches.flatMap((match, index) => match.pageKey === pageKey ? match.rects.map((rect, part) => ({ rect, index, part })) : [])
  if (hits.length === 0) return null
  return (
    <div className="pdf-find-layer">
      {hits.map(({ rect, index, part }) => (
        <div
          key={`${index}:${part}`}
          className={`pdf-find-hit${index === current ? ' is-current' : ''}`}
          style={{ left: rect.x * zoom, top: rect.y * zoom, width: rect.width * zoom, height: rect.height * zoom }}
        />
      ))}
    </div>
  )
}
