import { useEffect } from 'react'
import { CircleAlert, CircleArrowDown, X } from 'lucide-react'
import { toast } from 'sonner'
import { create } from 'zustand'
import { suiteName } from '@shared/applications'
import type { UpdateStatus } from '@shared/shell'
import { platformClient } from '../services/platform/client'

/** This window's copy of the main process's update status. */
const useUpdateStore = create<{ readonly status: UpdateStatus }>(() => ({ status: { state: 'none' } }))

function showOfferToast(): void {
  // One toast per window, replaced on the next check; it stays until closed so download progress stays visible.
  toast.custom(id => <UpdateOffer onClose={() => toast.dismiss(id)} />, { id: 'update', duration: Infinity })
}

/** Keeps the status current and shows the offer as a toast after each check that finds a release. */
export function useUpdateEvents(): void {
  useEffect(() => platformClient.onUpdateStatus(status => {
    useUpdateStore.setState({ status })
    // A later check found nothing to offer; the toast would be left empty.
    if (status.state === 'none') toast.dismiss('update')
  }), [])
  useEffect(() => platformClient.onUpdateOffered(showOfferToast), [])
}

/** Check for Updates: this window hears "up to date" or the error; a found release reaches every window as the offer. */
export async function checkForUpdates(): Promise<void> {
  const result = await platformClient.checkForUpdates()
  if (result.outcome === 'current') toast(`${suiteName} is up to date (${result.version})`)
  else if (result.outcome === 'failed') toast.error(result.message)
}

function describe(status: Exclude<UpdateStatus, { state: 'none' }>): string {
  switch (status.state) {
    case 'available': return `${suiteName} ${status.version} is available.`
    case 'downloading': return `Downloading… ${status.percent}%`
    case 'downloaded': return `Drag ${suiteName} into Applications and choose Replace, then reopen ${suiteName}.`
    case 'failed': return status.message
  }
}

/** The update offer and its actions: a banner on Home, and the toast when `onClose` is given. */
export function UpdateOffer({ onClose }: { readonly onClose?: () => void }) {
  const status = useUpdateStore(state => state.status)
  if (status.state === 'none') return null
  const Icon = status.state === 'failed' ? CircleAlert : CircleArrowDown
  return (
    <div className={`update-offer${onClose === undefined ? '' : ' is-toast'}`} role="status">
      <Icon aria-hidden="true" size={18} />
      <span className="update-offer-text">{describe(status)}</span>
      <span className="update-offer-actions">
        {(status.state === 'available' || status.state === 'failed') && (
          <button type="button" className="button is-primary" onClick={() => void platformClient.downloadUpdate()}>
            {status.state === 'failed' ? 'Try again' : 'Download'}
          </button>
        )}
        <button type="button" className="button is-quiet" onClick={() => void platformClient.openReleaseNotes()}>What’s new</button>
      </span>
      {onClose !== undefined && (
        <button type="button" className="tool update-offer-close" aria-label="Close" onClick={onClose}>
          <X aria-hidden="true" size={14} />
        </button>
      )}
    </div>
  )
}
