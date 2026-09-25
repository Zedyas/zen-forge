/**
 * Built-in colour themes, as plain data. A theme decides the colours and font that new slides,
 * text and shapes start with, and applying one recolours what still uses the old theme's colours.
 * Templates can later be added as a theme plus a set of starting slides.
 */
export interface Theme {
  readonly id: string
  readonly name: string
  readonly background: string
  /** Titles and body text. */
  readonly text: string
  /** Subtitles and secondary text. */
  readonly subtle: string
  /** New shapes' fill, and table header rows. */
  readonly accent: string
  /** Text on the accent colour. */
  readonly onAccent: string
  readonly font: string
}

/** White slides, dark Arial text: what every computer shows the same. */
export const defaultTheme: Theme = { id: 'light', name: 'Light', background: '#ffffff', text: '#222222', subtle: '#5f6368', accent: '#4a78c2', onAccent: '#ffffff', font: 'Arial' }

export const themes: readonly Theme[] = [
  defaultTheme,
  { id: 'paper', name: 'Paper', background: '#f7f4ec', text: '#2d2a26', subtle: '#6f6a60', accent: '#c2622d', onAccent: '#ffffff', font: 'Georgia' },
  { id: 'mint', name: 'Mint', background: '#eef7f2', text: '#173b2e', subtle: '#4f6f62', accent: '#2f8f6b', onAccent: '#ffffff', font: 'Arial' },
  { id: 'slate', name: 'Slate', background: '#1f2328', text: '#f2f3f5', subtle: '#aab1bb', accent: '#5b8def', onAccent: '#ffffff', font: 'Arial' },
  { id: 'ocean', name: 'Ocean', background: '#0f2a4a', text: '#ffffff', subtle: '#b8c7da', accent: '#e0a92c', onAccent: '#1b1b1b', font: 'Verdana' },
]

export function findTheme(id: string): Theme {
  return themes.find(theme => theme.id === id) ?? defaultTheme
}
