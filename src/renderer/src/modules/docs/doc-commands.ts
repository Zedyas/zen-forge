import { Extension, type Editor } from '@tiptap/core'
import { toast } from 'sonner'
import { create } from 'zustand'
import type { CommandId } from '@shared/commands'
import { fileService } from '../../services/file/IpcFileService'
import { platformClient } from '../../services/platform/client'
import { imageSize, mimeTypeForExtension, toDataUrl } from './io/images'
import { contentWidthPixels, documentPage, documentTheme, type BlockStyle, type PageSetup, type StyleName } from './theme'

/** The mounted document that menu commands act on. */
export interface DocController {
  readonly documentId: string
  readonly editor: Editor
  openFind(): void
  print(): Promise<void>
  exportPdf(): Promise<void>
}

let active: DocController | undefined

/** Registers the mounted editor as the target of menu commands; returns the unregisterer. */
export function setActiveDoc(controller: DocController): () => void {
  active = controller
  return () => {
    if (active === controller) active = undefined
  }
}

// ─── Zoom ───────────────────────────────────────────────────────────────

const zoomLevels: readonly number[] = [0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2]

/** Page zoom per open document; a view setting, so it is not saved or undone. */
export const useDocZoom = create<{ readonly zoom: Readonly<Record<string, number>> }>(() => ({ zoom: {} }))

function setZoom(documentId: string, step: 1 | -1 | 0): void {
  const current = useDocZoom.getState().zoom[documentId] ?? 1
  const next = step === 0 ? 1 : step === 1
    ? zoomLevels.find(level => level > current + 0.001)
    : [...zoomLevels].reverse().find(level => level < current - 0.001)
  if (next !== undefined) useDocZoom.setState(state => ({ zoom: { ...state.zoom, [documentId]: next } }))
}

// ─── Text and paragraph formatting ──────────────────────────────────────

/** The theme style of the block the caret is in: Title, a heading, Quote inside a quotation, else Normal. */
export function currentStyle(editor: Editor): StyleName {
  if (editor.isActive('title')) return 'title'
  for (const level of [1, 2, 3] as const) if (editor.isActive('heading', { level })) return `heading${level}`
  return editor.isActive('blockquote') ? 'quote' : 'normal'
}

export function currentBlockStyle(editor: Editor): BlockStyle {
  return documentTheme(editor.state.doc.attrs['theme'])[currentStyle(editor)]
}

export function setBlockStyle(editor: Editor, style: StyleName): void {
  const chain = editor.chain().focus()
  if (style === 'quote') {
    chain.setParagraph()
    if (!editor.isActive('blockquote')) chain.setBlockquote()
    chain.run()
    return
  }
  if (editor.isActive('blockquote')) chain.unsetBlockquote()
  if (style === 'normal') chain.setParagraph()
  else if (style === 'title') chain.setNode('title')
  else chain.setHeading({ level: style === 'heading1' ? 1 : style === 'heading2' ? 2 : 3 })
  chain.run()
}

export type ParagraphFormat = Partial<Record<'lineHeight' | 'spaceBefore' | 'spaceAfter' | 'indentLeft' | 'indentFirstLine' | 'textAlign', number | string | null>>

/** Sets direct paragraph formatting on every paragraph, heading and title in the selection. */
export function setParagraphFormat(editor: Editor, format: ParagraphFormat): void {
  editor.chain().focus()
    .updateAttributes('paragraph', format)
    .updateAttributes('heading', format)
    .updateAttributes('title', format)
    .run()
}

const indentStep = 36

/** In a list, moves the item a level in or out; elsewhere changes the left indent by half an inch. */
export function indent(editor: Editor, direction: 1 | -1): void {
  const item = editor.isActive('taskItem') ? 'taskItem' : editor.isActive('listItem') ? 'listItem' : undefined
  if (item !== undefined) {
    if (direction === 1) editor.chain().focus().sinkListItem(item).run()
    else editor.chain().focus().liftListItem(item).run()
    return
  }
  const { $from } = editor.state.selection
  const current: unknown = $from.parent.attrs['indentLeft']
  const base = currentBlockStyle(editor).indentLeft
  const next = Math.max(0, (typeof current === 'number' ? current : base) + direction * indentStep)
  setParagraphFormat(editor, { indentLeft: next === base ? null : next })
}

/** Removes direct formatting: marks other than links, and the paragraph's own spacing, indents and alignment. */
export function clearFormatting(editor: Editor): void {
  const chain = editor.chain().focus()
  for (const mark of ['bold', 'italic', 'underline', 'strike', 'code', 'superscript', 'subscript', 'textStyle']) chain.unsetMark(mark)
  chain.run()
  setParagraphFormat(editor, { lineHeight: null, spaceBefore: null, spaceAfter: null, indentLeft: null, indentFirstLine: null, textAlign: null })
}

/** Changes the page setup; one undo step, like any edit. */
export function setPage(editor: Editor, change: Partial<PageSetup>): void {
  const page = { ...documentPage(editor.state.doc.attrs['page']), ...change }
  editor.chain().command(({ tr }) => {
    tr.setDocAttribute('page', page)
    return true
  }).run()
}

// ─── Menu commands ──────────────────────────────────────────────────────

/** The native menu owns ⌘Z, so text fields (find, link, inspector) get their own undo back here. */
function isEditingField(): boolean {
  const focused = document.activeElement
  return focused instanceof HTMLInputElement || focused instanceof HTMLTextAreaElement
}

export function runDocCommand(command: CommandId): void | Promise<void> {
  if (active === undefined) return
  const { editor, documentId } = active
  switch (command) {
    case 'edit.undo':
      if (isEditingField()) document.execCommand('undo')
      else editor.chain().focus().undo().run()
      return
    case 'edit.redo':
      if (isEditingField()) document.execCommand('redo')
      else editor.chain().focus().redo().run()
      return
    case 'edit.find':
      return active.openFind()
    case 'format.bold':
      editor.chain().focus().toggleBold().run()
      return
    case 'format.italic':
      editor.chain().focus().toggleItalic().run()
      return
    case 'format.underline':
      editor.chain().focus().toggleUnderline().run()
      return
    case 'format.clear':
      return clearFormatting(editor)
    case 'view.zoomIn':
      return setZoom(documentId, 1)
    case 'view.zoomOut':
      return setZoom(documentId, -1)
    case 'view.actualSize':
      return setZoom(documentId, 0)
    case 'file.print':
      return active.print()
    case 'file.exportPdf':
      return active.exportPdf()
  }
}

// ─── Inserting ──────────────────────────────────────────────────────────

/** Adds a PNG or JPEG from disk at the caret. The image is stored in the document, as Word does. */
export async function insertImage(editor: Editor): Promise<void> {
  const [file] = await fileService.chooseFiles({ extensions: ['png', 'jpg', 'jpeg'] })
  if (file === undefined) return
  const bytes = await fileService.read(file.path)
  const src = toDataUrl(bytes, mimeTypeForExtension(file.extension))
  // The size it is placed at, as Word records it: its own size, scaled down to the text width.
  const natural = imageSize(bytes)
  const column = contentWidthPixels(documentPage(editor.state.doc.attrs['page']))
  const width = natural === undefined ? undefined : Math.round(Math.min(natural.width, column))
  const size = natural === undefined || width === undefined || natural.width === 0 ? {} : { width, height: Math.round(natural.height * (width / natural.width)) }
  editor.chain().focus().setImage({ src, ...size }).run()
}

/** A web or mail address as typed: `example.com` becomes `https://example.com`. Other schemes are refused. */
export function normalizeLink(input: string): string | undefined {
  const text = input.trim()
  if (text === '') return undefined
  if (/^(https?:\/\/|mailto:)/i.test(text)) return text
  if (/^[^\s@/]+@[^\s@/]+\.[^\s@/]+$/.test(text)) return `mailto:${text}`
  if (/^[a-z][a-z0-9+.-]*:/i.test(text) && !/^[^:/]+:\d/.test(text)) return undefined
  return `https://${text}`
}

/** Links the selection, or inserts the address as linked text when nothing is selected. */
export function applyLink(editor: Editor, input: string): boolean {
  const href = normalizeLink(input)
  if (href === undefined) return false
  if (editor.state.selection.empty && !editor.isActive('link')) {
    editor.chain().focus().insertContent({ type: 'text', text: href.replace(/^mailto:/, ''), marks: [{ type: 'link', attrs: { href } }] }).run()
  } else {
    editor.chain().focus().extendMarkRange('link').setLink({ href }).run()
  }
  return true
}

export function removeLink(editor: Editor): void {
  editor.chain().focus().extendMarkRange('link').unsetLink().run()
}

/** Opens a web link in the default browser. Other kinds of link never open; the app never navigates itself. */
export async function openLink(href: string): Promise<void> {
  if (!/^https?:\/\//i.test(href)) {
    toast.info('Only web links open from Zendo', { description: href })
    return
  }
  await platformClient.openExternal(href)
}

/**
 * Starts a new page at the caret, as Word does: at the start of a paragraph the break goes before
 * it, at the end it goes after it with an empty paragraph to type in, and in the middle it splits
 * the paragraph.
 */
export function insertPageBreak(editor: Editor): void {
  const { $from } = editor.state.selection
  const block = $from.parent
  if (!block.isTextblock || $from.depth === 0) {
    editor.chain().focus().insertContent({ type: 'pageBreak' }).run()
    return
  }
  if ($from.parentOffset === 0) {
    editor.chain().focus().insertContentAt($from.before(), { type: 'pageBreak' }).run()
    return
  }
  if ($from.parentOffset === block.content.size) {
    const after = $from.after()
    // The new paragraph starts after the break (one position) and its own opening (one more).
    editor.chain().focus().insertContentAt(after, [{ type: 'pageBreak' }, { type: 'paragraph' }]).setTextSelection(after + 2).run()
    return
  }
  editor.chain().focus().insertContent({ type: 'pageBreak' }).run()
}

/** Word's shortcuts for strikethrough and page break, ahead of TipTap's own (⌘↩ is a line break there). */
export const DocShortcuts = Extension.create({
  name: 'docShortcuts',
  priority: 1000,
  addKeyboardShortcuts() {
    return {
      'Mod-Shift-x': () => this.editor.commands.toggleStrike(),
      'Mod-Enter': () => {
        insertPageBreak(this.editor)
        return true
      },
    }
  },
})
