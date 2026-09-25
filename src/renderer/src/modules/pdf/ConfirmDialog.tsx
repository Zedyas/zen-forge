import { useEffect } from 'react'
import { Dialog } from '@base-ui/react/dialog'
import { create } from 'zustand'

interface ConfirmAction {
  readonly label: string
  readonly value: string
  readonly primary?: boolean
}

/** A checkbox shown above the buttons; its state comes back with the answer. */
interface ConfirmOption {
  readonly label: string
  readonly detail?: string
  readonly checked: boolean
}

export interface ConfirmAnswer {
  /** The chosen action's value, or `cancel`. */
  readonly value: string
  readonly optionChecked: boolean
}

interface ConfirmRequest {
  readonly title: string
  readonly message: string
  readonly actions: readonly ConfirmAction[]
  readonly option?: ConfirmOption
  readonly resolve: (answer: ConfirmAnswer) => void
}

const useConfirmStore = create<{ readonly request?: ConfirmRequest; readonly optionChecked: boolean }>(() => ({ optionChecked: false }))

/**
 * Shows a modal choice and resolves with the chosen action's value, or `cancel` when dismissed.
 * A question still open when a new one arrives, or when the dialog leaves the screen, is cancelled,
 * so no caller waits forever.
 */
export function askConfirm(options: Omit<ConfirmRequest, 'resolve'>): Promise<ConfirmAnswer> {
  return new Promise(resolve => {
    useConfirmStore.getState().request?.resolve({ value: 'cancel', optionChecked: false })
    useConfirmStore.setState({ request: { ...options, resolve }, optionChecked: options.option?.checked ?? false })
  })
}

function answer(value: string): void {
  const { request, optionChecked } = useConfirmStore.getState()
  useConfirmStore.setState({ request: undefined })
  request?.resolve({ value, optionChecked })
}

export function ConfirmHost() {
  const request = useConfirmStore(state => state.request)
  const optionChecked = useConfirmStore(state => state.optionChecked)
  useEffect(() => () => answer('cancel'), [])
  return (
    <Dialog.Root open={request !== undefined} onOpenChange={open => { if (!open) answer('cancel') }}>
      <Dialog.Portal>
        <Dialog.Backdrop className="pdf-dialog-backdrop" />
        <Dialog.Popup className="pdf-dialog">
          <Dialog.Title render={<h2 />}>{request?.title}</Dialog.Title>
          <Dialog.Description>{request?.message}</Dialog.Description>
          {request?.option !== undefined && (
            <label className="pdf-dialog-option">
              <input
                type="checkbox"
                checked={optionChecked}
                onChange={event => useConfirmStore.setState({ optionChecked: event.currentTarget.checked })}
              />
              <span>
                {request.option.label}
                {request.option.detail !== undefined && <small>{request.option.detail}</small>}
              </span>
            </label>
          )}
          <div className="pdf-dialog-actions">
            <button type="button" className="button" onClick={() => answer('cancel')}>Cancel</button>
            {request?.actions.map(action => (
              <button
                key={action.value}
                type="button"
                className={`button${action.primary === true ? ' is-primary' : ''}`}
                onClick={() => answer(action.value)}
              >
                {action.label}
              </button>
            ))}
          </div>
        </Dialog.Popup>
      </Dialog.Portal>
    </Dialog.Root>
  )
}
