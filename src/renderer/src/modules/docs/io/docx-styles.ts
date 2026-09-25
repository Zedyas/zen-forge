/** Names the .docx writer and reader share, kept apart so reading does not load the writer. */

/** Styles for what Word has no built-in style for, by id; the reader maps them back by name. */
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
