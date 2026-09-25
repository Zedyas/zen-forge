import { Popover } from '@base-ui/react/popover'
import { ChevronDown, type LucideIcon } from 'lucide-react'
import type { ReactNode } from 'react'
import type { CommandId } from '@shared/commands'
import { commandShortcut, Tip } from './Tip'

interface ToolButtonProps {
  readonly icon?: LucideIcon
  readonly label: string
  /** Shown instead of the icon, for tools whose clearest symbol is text (".00", "123"). */
  readonly text?: string
  /** Adds the command's shortcut to the tooltip. */
  readonly command?: CommandId
  /** A shortcut that is not a menu command, such as a single-key tool switch. */
  readonly shortcut?: string
  readonly pressed?: boolean
  readonly disabled?: boolean
  onClick(): void
}

export function ToolButton({ icon: Icon, label, text, command, shortcut, pressed, disabled, onClick }: ToolButtonProps) {
  return (
    <Tip label={label} shortcut={shortcut ?? commandShortcut(command)}>
      <button
        type="button"
        className={`tool${text !== undefined ? ' is-label' : ''}`}
        aria-label={label}
        aria-pressed={pressed}
        disabled={disabled}
        onClick={onClick}
      >
        {Icon !== undefined && <Icon aria-hidden="true" size={16} strokeWidth={1.7} />}
        {text}
      </button>
    </Tip>
  )
}

export function ToolSeparator() {
  return <span className="tool-sep" aria-hidden="true" />
}

export interface ColorOption {
  readonly label: string
  /** `undefined` means no colour: the default text colour or no fill. */
  readonly value: string | undefined
}

interface ColorToolProps {
  readonly icon: LucideIcon
  readonly label: string
  readonly options: readonly ColorOption[]
  /** The colour the main button applies; shown as the bar under the icon. */
  readonly current: string | undefined
  readonly disabled?: boolean
  onApply(value: string | undefined): void
  onChoose(value: string | undefined): void
}

/** Split button: the icon applies the last chosen colour in one click; the caret picks another. */
export function ColorTool({ icon: Icon, label, options, current, disabled, onApply, onChoose }: ColorToolProps) {
  return (
    <span className="split-tool">
      <Tip label={label}>
        <button type="button" className="tool" aria-label={label} disabled={disabled} onClick={() => onApply(current)}>
          <Icon aria-hidden="true" size={16} strokeWidth={1.7} />
          <i className="swatch-bar" style={{ background: current ?? 'transparent' }} />
        </button>
      </Tip>
      <Popover.Root>
        <Tip label={`Choose ${label.toLowerCase()}`}>
          <Popover.Trigger className="tool caret" aria-label={`Choose ${label.toLowerCase()}`} disabled={disabled}>
            <ChevronDown aria-hidden="true" size={11} strokeWidth={2.2} />
          </Popover.Trigger>
        </Tip>
        <Popover.Portal>
          <Popover.Positioner sideOffset={6} align="start">
            <Popover.Popup className="popover">
              <div className="popover-label">{label}</div>
              <div className="swatch-grid">
                {options.map(option => (
                  <Tip key={option.label} label={option.label}>
                    <Popover.Close
                      className={`swatch${option.value === undefined ? ' is-none' : ''}`}
                      style={option.value === undefined ? undefined : { background: option.value }}
                      aria-label={option.label}
                      aria-pressed={option.value === current}
                      onClick={() => onChoose(option.value)}
                    />
                  </Tip>
                ))}
              </div>
            </Popover.Popup>
          </Popover.Positioner>
        </Popover.Portal>
      </Popover.Root>
    </span>
  )
}

export function Toolbar({ label, children }: { readonly label: string; readonly children: ReactNode }) {
  return <div className="toolbar" role="toolbar" aria-label={label}>{children}</div>
}
