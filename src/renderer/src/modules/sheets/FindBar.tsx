import { CaseSensitive, ChevronDown, ChevronUp, X } from 'lucide-react'
import { useMemo, useState } from 'react'
import { toast } from 'sonner'
import type { FindMatch, Workbook } from './model/Workbook'
import { ToolButton } from '../../ui/Toolbar'

interface FindBarProps {
  readonly workbook: Workbook
  readonly revision: number
  onClose(): void
  onSelectMatch(match: FindMatch): void
}

/** Find and replace across every sheet, searching cell contents including formula text. */
export function FindBar({ workbook, revision, onClose, onSelectMatch }: FindBarProps) {
  const [query, setQuery] = useState('')
  const [replacement, setReplacement] = useState('')
  const [caseSensitive, setCaseSensitive] = useState(false)
  const [matchIndex, setMatchIndex] = useState(0)
  // `revision` re-runs the search after edits; the workbook object itself does not change identity.
  const matches = useMemo(() => workbook.find(query, caseSensitive), [caseSensitive, query, revision, workbook])
  const current = matches.length === 0 ? 0 : matchIndex % matches.length

  const move = (difference: -1 | 1): void => {
    if (matches.length === 0) return
    const next = (current + difference + matches.length) % matches.length
    setMatchIndex(next)
    const match = matches[next]
    if (match !== undefined) onSelectMatch(match)
  }

  return (
    <div className="find-bar" role="search">
      <input
        autoFocus
        className="text-field"
        value={query}
        placeholder="Find"
        aria-label="Find"
        onChange={event => {
          setQuery(event.currentTarget.value)
          setMatchIndex(0)
        }}
        onKeyDown={event => {
          if (event.key === 'Enter') move(event.shiftKey ? -1 : 1)
          if (event.key === 'Escape') onClose()
        }}
      />
      <ToolButton icon={CaseSensitive} label="Match case" pressed={caseSensitive} onClick={() => setCaseSensitive(value => !value)} />
      <output className="find-count">{query === '' ? '' : matches.length === 0 ? 'No matches' : `${current + 1} of ${matches.length}`}</output>
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
          if (event.key === 'Escape') onClose()
        }}
      />
      <button
        type="button"
        className="button"
        disabled={matches.length === 0}
        onClick={() => {
          const replaced = workbook.replaceAll(query, replacement, caseSensitive)
          toast.success(`Replaced ${replaced} ${replaced === 1 ? 'cell' : 'cells'}`)
        }}
      >
        Replace all
      </button>
      <span className="toolbar-grow" />
      <ToolButton icon={X} label="Close find" onClick={onClose} />
    </div>
  )
}
