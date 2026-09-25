import type { HorizontalAlign, ListStyle } from './model'

/** Formatting applied to the selected text while a text box is being typed into. */
export type InlineFormat =
  | { readonly kind: 'toggle'; readonly style: 'bold' | 'italic' | 'underline' }
  | { readonly kind: 'color'; readonly value: string }
  | { readonly kind: 'highlight'; readonly value: string | undefined }
  | { readonly kind: 'size'; readonly value: number }
  | { readonly kind: 'font'; readonly value: string }
  | { readonly kind: 'align'; readonly value: HorizontalAlign }
  | { readonly kind: 'list'; readonly value: Exclude<ListStyle, 'none'> }
  | { readonly kind: 'indent'; readonly step: 1 | -1 }

/** The open inline text editor. Toolbar and menu formatting goes to its selection instead of the whole box. */
export interface ActiveEditor {
  readonly documentId: string
  readonly elementId: string
  format(change: InlineFormat): void
  /** Writes the typed text into the presentation and closes the editor. */
  commit(): void
}

let active: ActiveEditor | undefined

export function activeEditor(documentId: string): ActiveEditor | undefined {
  return active?.documentId === documentId ? active : undefined
}

/** Registers an editor while it is open; returns the function that unregisters it. */
export function registerEditor(editor: ActiveEditor): () => void {
  active = editor
  return () => {
    if (active === editor) active = undefined
  }
}
