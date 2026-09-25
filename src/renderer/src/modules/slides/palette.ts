/** Colour choices offered by the toolbar and inspector. */

export const textColors: ReadonlyArray<{ readonly label: string; readonly value: string }> = [
  { label: 'Dark', value: '#222222' },
  { label: 'Grey', value: '#5f6368' },
  { label: 'White', value: '#ffffff' },
  { label: 'Red', value: '#c9352b' },
  { label: 'Orange', value: '#d9730d' },
  { label: 'Green', value: '#1f8a4c' },
  { label: 'Blue', value: '#2f5bd3' },
  { label: 'Purple', value: '#7a3fc8' },
]

export const fillColors: ReadonlyArray<{ readonly label: string; readonly value: string | undefined }> = [
  { label: 'No fill', value: undefined },
  { label: 'Blue', value: '#4a78c2' },
  { label: 'Green', value: '#2fb592' },
  { label: 'Yellow', value: '#e0a92c' },
  { label: 'Red', value: '#e05a47' },
  { label: 'Light grey', value: '#e9ebee' },
  { label: 'Dark', value: '#222222' },
  { label: 'White', value: '#ffffff' },
]

export const highlightColors: ReadonlyArray<{ readonly label: string; readonly value: string | undefined }> = [
  { label: 'No highlight', value: undefined },
  { label: 'Yellow', value: '#fff176' },
  { label: 'Green', value: '#b9f6ca' },
  { label: 'Blue', value: '#b3e5fc' },
  { label: 'Pink', value: '#f8bbd0' },
  { label: 'Orange', value: '#ffe0b2' },
  { label: 'Purple', value: '#e1bee7' },
  { label: 'Grey', value: '#e0e0e0' },
]

export const backgroundColors: ReadonlyArray<{ readonly label: string; readonly value: string }> = [
  { label: 'White', value: '#ffffff' },
  { label: 'Paper', value: '#f7f4ec' },
  { label: 'Light grey', value: '#eef0f3' },
  { label: 'Light blue', value: '#e3ecfa' },
  { label: 'Light green', value: '#e2f4ea' },
  { label: 'Charcoal', value: '#2b2d31' },
  { label: 'Navy', value: '#1d2b4f' },
  { label: 'Black', value: '#000000' },
]

/** The fonts offered everywhere; a font found in an opened file is offered too. */
export const fonts: readonly string[] = ['Arial', 'Helvetica', 'Georgia', 'Times New Roman', 'Verdana', 'Courier New']

export function fontChoices(current: string | undefined): readonly string[] {
  return current === undefined || fonts.includes(current) ? fonts : [...fonts, current]
}

export const fontSizes: readonly number[] = [10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 72, 96]

export const lineSpacings: readonly number[] = [1, 1.15, 1.5, 2, 2.5, 3]
