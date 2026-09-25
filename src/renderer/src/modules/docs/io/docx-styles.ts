/** Names and colours the .docx writer and reader share, kept apart so reading does not load the writer. */

/** Paragraph and character styles for what Word has no built-in style for, by id; the reader maps them back by name. */
export const docxStyleNames = {
  Quote: 'Quote',
  Code: 'Code',
  InlineCode: 'Inline Code',
  HorizontalLine: 'Horizontal Line',
  ListContinue: 'List Continue',
  Checklist: 'Checklist',
  TableText: 'Table Text',
} as const

export type DocxStyleId = keyof typeof docxStyleNames

/** Word has no checklist; a checklist item is a paragraph that starts with one of these boxes. */
export const checklistMarks = { open: '☐', done: '☑' } as const

/** Word's highlight colours (`w:highlight`) by name. Any other background colour is run shading. */
export const highlightColors: Readonly<Record<string, string>> = {
  yellow: '#ffff00',
  green: '#00ff00',
  cyan: '#00ffff',
  magenta: '#ff00ff',
  blue: '#0000ff',
  red: '#ff0000',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#c0c0c0',
  black: '#000000',
  white: '#ffffff',
}

/** `#abc`, `#aabbcc` or `rgb(…)` as lowercase `#aabbcc`; undefined for anything else. */
export function hexColor(value: string): string | undefined {
  const text = value.trim().toLowerCase()
  const short = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(text)
  if (short !== null) return `#${short[1]}${short[1]}${short[2]}${short[2]}${short[3]}${short[3]}`
  if (/^#[0-9a-f]{6}$/.test(text)) return text
  const rgb = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(text)
  if (rgb === null) return undefined
  return `#${rgb.slice(1, 4).map(part => Number(part).toString(16).padStart(2, '0')).join('')}`
}
