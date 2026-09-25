import { Tooltip } from '@base-ui/react/tooltip'
import type { ReactElement, ReactNode } from 'react'
import { acceleratorLabel, findCommand, type CommandId } from '@shared/commands'

/** The menu-style shortcut for a command ("⌘B"), or undefined when it has none. */
export function commandShortcut(command: CommandId | undefined): string | undefined {
  const accelerator = command === undefined ? undefined : findCommand(command).accelerator
  return accelerator === undefined ? undefined : acceleratorLabel(accelerator)
}

interface TipProps {
  readonly label: string
  readonly shortcut?: string | undefined
  /** One element that forwards props and ref: a <button>, or a Base UI trigger such as Popover.Trigger. */
  readonly children: ReactElement
}

/** Hover label for a control: its name and, when it has one, the keyboard shortcut. */
export function Tip({ label, shortcut, children }: TipProps) {
  return (
    <Tooltip.Root disableHoverablePopup>
      <Tooltip.Trigger render={children} />
      <Tooltip.Portal>
        <Tooltip.Positioner className="tip-positioner" side="bottom" sideOffset={6}>
          <Tooltip.Popup className="tip">
            {label}
            {shortcut !== undefined && <kbd>{shortcut}</kbd>}
          </Tooltip.Popup>
        </Tooltip.Positioner>
      </Tooltip.Portal>
    </Tooltip.Root>
  )
}

/**
 * Shares one hover delay across the window: the first tooltip waits, then moving along a toolbar
 * shows the next one immediately.
 */
export function TipProvider({ children }: { readonly children: ReactNode }) {
  return <Tooltip.Provider delay={450} closeDelay={0}>{children}</Tooltip.Provider>
}
