import { CaseSensitive, ChevronDown, ChevronUp, X } from 'lucide-react'
import { useEffect, useMemo, useRef, useState } from 'react'
import { toast } from 'sonner'
import { ToolButton } from '../../ui/Toolbar'
import { markMatch, selectInNotes } from './find-reveal'
import { closeFind, replaceFound, revealMatch } from './slides-actions'
import { findMatches, occurrences, type FindMatch } from './slides-find'
import { readySlides, type ReadySlides } from './slides-store'

interface FindBarProps {
  readonly documentId: string
  readonly document: ReadySlides
  /** Changes on every ⌘F, which focuses the find field again. */
  readonly focusRequest: number
}

/** Find and replace across every slide's text boxes, shapes, table cells and speaker notes. */
export function FindBar({ documentId, document, focusRequest }: FindBarProps) {
  const [query, setQuery] = useState('')
  const [replacement, setReplacement] = useState('')
  const [caseSensitive, setCaseSensitive] = useState(false)
  // The match last shown, or -1 before any: the first Return shows the first match.
  const [shown, setShown] = useState(-1)
  // Counts the steps to a match, so stepping to the same one again (the only match) selects it again.
  const [reveals, setReveals] = useState(0)
  const input = useRef<HTMLInputElement>(null)
  const matches = useMemo(() => findMatches(document.present, query, caseSensitive), [caseSensitive, document.present, query])
  const current = matches.length === 0 ? 0 : Math.max(shown, 0) % matches.length

  useEffect(() => {
    input.current?.focus()
    input.current?.select()
  }, [focusRequest])

  // Marks the shown occurrence, so stepping between two matches in one text box visibly moves.
  const shownMatch = shown < 0 ? undefined : matches[current]
  useEffect(() => (shownMatch === undefined ? undefined : markMatch(shownMatch)), [shownMatch])
  // Only on a step: the notes field takes focus, which neither an edit of the presentation nor Replace may do.
  useEffect(() => {
    if (shownMatch !== undefined) selectInNotes(shownMatch)
  }, [reveals])

  const show = (index: number, list: readonly FindMatch[]): void => {
    const match = list[index]
    if (match === undefined) return
    setShown(index)
    revealMatch(documentId, match)
  }

  const move = (step: -1 | 1): void => {
    const from = shown < 0 ? (step === 1 ? -1 : 0) : current
    show((from + step + matches.length) % matches.length, matches)
    setReveals(count => count + 1)
  }

  const replaceCurrent = (): void => {
    const match = matches[current]
    if (match === undefined) return
    replaceFound(documentId, [match], replacement)
    // The next match takes the replaced one's place, after any the replacement text itself holds.
    const remaining = findMatches(readySlides(documentId)?.present ?? document.present, query, caseSensitive)
    if (remaining.length === 0) setShown(-1)
    else show((current + occurrences(replacement, query, caseSensitive)) % remaining.length, remaining)
  }

  const replaceAll = (): void => {
    replaceFound(documentId, matches, replacement)
    setShown(-1)
    toast.success(`Replaced ${matches.length} ${matches.length === 1 ? 'match' : 'matches'}`)
  }

  const count = query === '' ? '' : matches.length === 0 ? 'No matches' : shown < 0 ? `${matches.length} ${matches.length === 1 ? 'match' : 'matches'}` : `${current + 1} of ${matches.length}`
  const inNotes = shown >= 0 && matches[current]?.where === 'notes'

  return (
    <div className="find-bar" role="search">
      <input
        ref={input}
        className="text-field"
        value={query}
        placeholder="Find"
        aria-label="Find"
        onChange={event => {
          setQuery(event.currentTarget.value)
          setShown(-1)
        }}
        onKeyDown={event => {
          if (event.key === 'Enter') move(event.shiftKey ? -1 : 1)
          if (event.key === 'Escape') closeFind(documentId)
        }}
      />
      <ToolButton icon={CaseSensitive} label="Match case" pressed={caseSensitive} onClick={() => { setCaseSensitive(value => !value); setShown(-1) }} />
      <output className="find-count">{inNotes ? `${count} · in notes` : count}</output>
      <ToolButton icon={ChevronUp} label="Previous match" disabled={matches.length === 0} onClick={() => move(-1)} />
      <ToolButton icon={ChevronDown} label="Next match" disabled={matches.length === 0} onClick={() => move(1)} />
      <span className="tool-sep" />
      <input
        className="text-field"
        value={replacement}
        placeholder="Replace with"
        aria-label="Replace with"
        onChange={event => setReplacement(event.currentTarget.value)}
        onKeyDown={event => {
          if (event.key === 'Enter') replaceCurrent()
          if (event.key === 'Escape') closeFind(documentId)
        }}
      />
      <button type="button" className="button" disabled={matches.length === 0} onClick={replaceCurrent}>Replace</button>
      <button type="button" className="button" disabled={matches.length === 0} onClick={replaceAll}>Replace all</button>
      <span className="toolbar-grow" />
      <ToolButton icon={X} label="Close find" onClick={() => closeFind(documentId)} />
    </div>
  )
}
