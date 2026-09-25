import { create } from 'zustand'

interface ShellUiState {
  readonly paletteOpen: boolean
  setPaletteOpen(open: boolean): void
}

export const useShellUiStore = create<ShellUiState>(set => ({
  paletteOpen: false,
  setPaletteOpen: paletteOpen => set({ paletteOpen }),
}))
