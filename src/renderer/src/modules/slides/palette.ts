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
