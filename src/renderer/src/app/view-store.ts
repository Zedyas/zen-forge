import { create } from 'zustand'
import type { ViewState } from '@shared/shell'
import { platformClient } from '../services/platform/client'

interface ViewStore extends ViewState {
  toggle(part: keyof ViewState): void
}

const storageKey = 'view'

function readViewState(): ViewState {
  try {
    const stored: unknown = JSON.parse(localStorage.getItem(storageKey) ?? 'null')
    if (typeof stored === 'object' && stored !== null && 'toolbar' in stored && 'inspector' in stored) {
      return { toolbar: stored.toolbar !== false, inspector: stored.inspector === true }
    }
  } catch {
    // Unreadable view state falls back to the default layout.
  }
  return { toolbar: true, inspector: false }
}

/** Toolbar and inspector visibility, shared by every tab, mirrored to the View menu check marks. */
export const useViewStore = create<ViewStore>((set, get) => ({
  ...readViewState(),
  toggle: part => {
    set({ [part]: !get()[part] })
    const { toolbar, inspector } = get()
    try {
      localStorage.setItem(storageKey, JSON.stringify({ toolbar, inspector }))
    } catch {
      // Not persisting only costs the preference on the next launch.
    }
    void platformClient.setViewState({ toolbar, inspector })
  },
}))
