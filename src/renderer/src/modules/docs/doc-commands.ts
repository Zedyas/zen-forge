import { Extension, type Editor } from '@tiptap/core'
import { toast } from 'sonner'
import type { CommandId } from '@shared/commands'
import { fileService } from '../../services/file/IpcFileService'
import { platformClient } from '../../services/platform/client'
import { mimeTypeForExtension, toDataUrl } from './io/images'

/** The mounted document that menu commands act on. */
export interface DocController {
  readonly documentId: string
  readonly editor: Editor
  openFind(): void
}

let active: DocController | undefined

/** Registers the mounted editor as the target of Edit and Format menu commands; returns the unregisterer. */
export function setActiveDoc(controller: DocController): () => void {
  active = controller
  return () => {
    if (active === controller) active = undefined
  }
}

/** The native menu owns ⌘Z, so text fields (find, link) get their own undo back here. */
function isEditingField(): boolean {
  const focused = document.activeElement
  return focused instanceof HTMLInputElement || focused instanceof HTMLTextAreaElement
}

/** Removes character formatting and alignment; links, headings and lists stay. */
export function clearFormatting(editor: Editor): void {
  editor.chain().focus()
    .unsetMark('bold').unsetMark('italic').unsetMark('underline').unsetMark('strike').unsetMark('code')
    .unsetTextAlign()
    .run()
}

export function runDocCommand(command: CommandId): void {
  if (active === undefined) return
  const { editor } = active
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
  }
}

/** Adds a PNG or JPEG from disk at the caret. The image is stored in the document, as Word does. */
export async function insertImage(editor: Editor): Promise<void> {
  const [file] = await fileService.chooseFiles({ extensions: ['png', 'jpg', 'jpeg'] })
  if (file === undefined) return
  const src = toDataUrl(await fileService.read(file.path), mimeTypeForExtension(file.extension))
  editor.chain().focus().setImage({ src }).run()
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

/** Opens a link in the default browser. Only https links open; the app never navigates itself. */
export async function openLink(href: string): Promise<void> {
  if (!href.startsWith('https://')) {
    toast.info('Only https links open from Zendo', { description: href })
    return
  }
  await platformClient.openExternal(href)
}

export function insertPageBreak(editor: Editor): void {
  const { $from } = editor.state.selection
  // In an empty paragraph the break takes its place; elsewhere it goes after the current block,
  // with an empty paragraph after it to keep typing on the new page.
  const emptyParagraph = $from.parent.type.name === 'paragraph' && $from.parent.content.size === 0
  editor.chain().focus().insertContent(emptyParagraph ? [{ type: 'pageBreak' }, { type: 'paragraph' }] : { type: 'pageBreak' }).run()
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
