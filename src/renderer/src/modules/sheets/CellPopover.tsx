import { Popover } from '@base-ui/react/popover'
import type { LucideIcon } from 'lucide-react'
import { useState, type FormEvent, type ReactNode } from 'react'
import { icon } from '../../ui/icons'
import { Tip } from '../../ui/Tip'

interface CellPopoverProps {
  readonly icon: LucideIcon
  readonly label: string
  readonly pressed?: boolean
  /** The field's text each time the popover opens. */
  readonly initial: string
  readonly placeholder: string
  /** A text area (Return adds a line, ⌘Return saves) instead of a one-line field. */
  readonly multiline?: boolean
  readonly submitLabel: string
  /** Applies the text; returning false keeps the popover open, after the caller has said why. */
  onSubmit(text: string): boolean
  /** Shown as a Remove button when set. */
  readonly onRemove?: () => void
  /** Below the field: a preview, examples… `setDraft` replaces the field's text. */
  readonly children?: (draft: string, setDraft: (text: string) => void) => ReactNode
}

/** A toolbar button whose popover edits one text property of the selected cells: a format code, a link, a note. */
export function CellPopover({ icon: Icon, label, pressed, initial, placeholder, multiline, submitLabel, onSubmit, onRemove, children }: CellPopoverProps) {
  const [open, setOpen] = useState(false)
  const [draft, setDraft] = useState('')

  const submit = (event?: FormEvent): void => {
    event?.preventDefault()
    if (onSubmit(draft)) setOpen(false)
  }

  return (
    <Popover.Root
      open={open}
      onOpenChange={next => {
        if (next) setDraft(initial)
        setOpen(next)
      }}
    >
      <Tip label={label}>
        <Popover.Trigger className="tool" aria-label={label} aria-pressed={pressed}>
          <Icon {...icon} />
        </Popover.Trigger>
      </Tip>
      <Popover.Portal>
        <Popover.Positioner sideOffset={6} align="start">
          <Popover.Popup className="popover sheet-cell-popover">
            <div className="popover-label">{label}</div>
            <form onSubmit={submit}>
              {multiline === true
                ? (
                  <textarea
                    className="text-field"
                    value={draft}
                    rows={4}
                    placeholder={placeholder}
                    aria-label={label}
                    onChange={event => setDraft(event.currentTarget.value)}
                    onKeyDown={event => {
                      if (event.key === 'Enter' && event.metaKey) submit()
                    }}
                  />
                )
                : (
                  <input
                    className="text-field"
                    value={draft}
                    placeholder={placeholder}
                    aria-label={label}
                    spellCheck={false}
                    onChange={event => setDraft(event.currentTarget.value)}
                  />
                )}
              {children?.(draft, setDraft)}
              <div className="sheet-cell-popover-actions">
                {onRemove !== undefined && (
                  <button type="button" className="button" onClick={() => {
                    onRemove()
                    setOpen(false)
                  }}>Remove</button>
                )}
                <button type="submit" className="button is-primary">{submitLabel}</button>
              </div>
            </form>
          </Popover.Popup>
        </Popover.Positioner>
      </Popover.Portal>
    </Popover.Root>
  )
}
