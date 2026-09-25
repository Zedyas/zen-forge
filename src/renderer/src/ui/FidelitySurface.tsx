import { Dialog } from '@base-ui/react/dialog'
import { ShieldAlert, X } from 'lucide-react'
import type { FindingSeverity } from '@shared/fidelity'
import { useFidelityStore } from '../services/fidelity/fidelity-store'
import { smallIcon } from './icons'
import { Tip } from './Tip'

const severityLabel: Record<FindingSeverity, string> = {
  degraded: 'Approximated',
  dropped: 'Lost on save',
}

/**
 * The status-bar entry for the active document's import report, its detail panel, and the
 * blocking warning shown when saving would remove content.
 */
export function FidelitySurface({ documentId }: { readonly documentId: string | undefined }) {
  const report = useFidelityStore(state => documentId === undefined ? undefined : state.reports[documentId])
  const panelOpen = useFidelityStore(state => state.panelOpen)
  const modalOpen = useFidelityStore(state => documentId !== undefined && state.modalDocumentId === documentId)
  const setPanelOpen = useFidelityStore(state => state.setPanelOpen)
  const closeModal = useFidelityStore(state => state.closeModal)

  // A file opened with nothing lost needs no mention.
  if (report === undefined || report.findings.length === 0) return null
  const dropped = report.findings.filter(finding => finding.severity === 'dropped')

  return (
    <>
      <button className="fidelity-status" type="button" onClick={() => setPanelOpen(!panelOpen)}>
        <i className={`fidelity-badge fidelity-badge--${report.severity}`} aria-hidden="true" />
        {report.findings.length} import {report.findings.length === 1 ? 'note' : 'notes'}
      </button>

      {panelOpen && (
        <aside className="fidelity-panel" aria-label="Import report">
          <header>
            <span>Import report</span>
            <Tip label="Close import report">
              <button type="button" className="tool" onClick={() => setPanelOpen(false)} aria-label="Close import report">
                <X {...smallIcon} />
              </button>
            </Tip>
          </header>
          <p>{report.sourceName} uses features Zendo does not represent exactly.</p>
          <ul>
            {report.findings.map((finding, index) => (
              <li key={`${finding.construct}-${index}`} data-severity={finding.severity}>
                <span>
                  <strong>{finding.construct}</strong>
                  {finding.location !== undefined && <small>{finding.location}</small>}
                  {finding.suggestedAlternative !== undefined && <small>{finding.suggestedAlternative}</small>}
                </span>
                <em>{severityLabel[finding.severity]}</em>
              </li>
            ))}
          </ul>
        </aside>
      )}

      <Dialog.Root open={modalOpen} onOpenChange={open => { if (!open) closeModal() }}>
        <Dialog.Portal>
          <Dialog.Backdrop className="fidelity-modal-backdrop" />
          <Dialog.Popup className="fidelity-modal">
            <ShieldAlert aria-hidden="true" size={26} />
            <Dialog.Title render={<h2 />}>Some of this file can't be kept</Dialog.Title>
            <Dialog.Description>
              You can view and edit {report.sourceName}. Saving suggests a new copy, so the original can keep everything. A file Zendo saves won't include:
            </Dialog.Description>
            <ul>{dropped.map(finding => <li key={finding.construct}>{finding.construct}</li>)}</ul>
            <div className="fidelity-modal-actions">
              <button type="button" className="button" onClick={() => {
                closeModal()
                setPanelOpen(true)
              }}>Show import report</button>
              <Dialog.Close className="button is-primary">Continue</Dialog.Close>
            </div>
          </Dialog.Popup>
        </Dialog.Portal>
      </Dialog.Root>
    </>
  )
}
