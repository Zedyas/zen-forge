import { useState, type FormEvent } from 'react'
import { useSortable } from '@dnd-kit/sortable'
import { X } from 'lucide-react'
import type { WorkbookSheet } from './model/Workbook'
import { tinyIcon } from '../../ui/icons'

interface SheetTabProps {
  readonly sheet: WorkbookSheet
  readonly active: boolean
  readonly canDelete: boolean
  onSelect(): void
  onRename(name: string): boolean
  onDelete(): void
}

export function SheetTab({
  sheet,
  active,
  canDelete,
  onSelect,
  onRename,
  onDelete,
}: SheetTabProps) {
  const [renaming, setRenaming] = useState(false)
  const [draft, setDraft] = useState(sheet.name)
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({ id: sheet.id })
  const style = transform === null
    ? undefined
    : { transform: `translate3d(${transform.x}px, ${transform.y}px, 0)`, transition }

  const commitRename = (event?: FormEvent): void => {
    event?.preventDefault()
    const name = draft.trim()
    if (name.length > 0 && name !== sheet.name) {
      if (!onRename(name)) setDraft(sheet.name)
    } else setDraft(sheet.name)
    setRenaming(false)
  }

  return (
    <div
      ref={setNodeRef}
      style={style}
      className={`sheet-tab${active ? ' is-active' : ''}${isDragging ? ' is-dragging' : ''}`}
    >
      {renaming ? (
        <form onSubmit={commitRename}>
          <input
            autoFocus
            value={draft}
            onChange={event => setDraft(event.currentTarget.value)}
            onBlur={() => commitRename()}
            onKeyDown={event => {
              if (event.key === 'Escape') {
                setDraft(sheet.name)
                setRenaming(false)
              }
            }}
            aria-label="Sheet name"
          />
        </form>
      ) : (
        <button
          type="button"
          {...attributes}
          {...listeners}
          role="tab"
          aria-selected={active}
          onClick={onSelect}
          onDoubleClick={() => setRenaming(true)}
        >
          {sheet.name}
        </button>
      )}
      {active && canDelete && !renaming && (
        <button
          className="sheet-delete"
          type="button"
          aria-label={`Delete ${sheet.name}`}
          // Deleting is undoable with ⌘Z, so it does not ask first.
          onClick={onDelete}
        >
          <X {...tinyIcon} />
        </button>
      )}
    </div>
  )
}
