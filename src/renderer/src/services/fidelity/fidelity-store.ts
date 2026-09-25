import { create } from 'zustand'
import type { ImportReport } from '@shared/fidelity'

interface FidelityState {
  /** Import reports by open document id. */
  readonly reports: Readonly<Record<string, ImportReport>>
  readonly panelOpen: boolean
  /** The document whose dropped-content warning is showing. */
  readonly modalDocumentId?: string
  publish(documentId: string, report: ImportReport): void
  forget(documentId: string): void
  setPanelOpen(open: boolean): void
  closeModal(): void
}

export const useFidelityStore = create<FidelityState>(set => ({
  reports: {},
  panelOpen: false,
  publish: (documentId, report) => set(state => ({
    reports: { ...state.reports, [documentId]: report },
    panelOpen: report.severity === 'degraded' ? true : state.panelOpen,
    modalDocumentId: report.severity === 'dropped' ? documentId : state.modalDocumentId,
  })),
  forget: documentId => set(state => {
    const reports = { ...state.reports }
    delete reports[documentId]
    return { reports, modalDocumentId: state.modalDocumentId === documentId ? undefined : state.modalDocumentId }
  }),
  setPanelOpen: panelOpen => set({ panelOpen }),
  closeModal: () => set({ modalDocumentId: undefined }),
}))
