import { mergeAttributes, Node, type Extensions } from '@tiptap/core'
import Image from '@tiptap/extension-image'
import { TaskItem, TaskList } from '@tiptap/extension-list'
import { TableKit } from '@tiptap/extension-table'
import TextAlign from '@tiptap/extension-text-align'
import StarterKit from '@tiptap/starter-kit'

/**
 * A manual page break. The editor flows continuously, so it shows as a labelled rule; Word and
 * print start a new page there. Markdown has no page break; the Markdown writer leaves it out.
 */
export const PageBreak = Node.create({
  name: 'pageBreak',
  group: 'block',
  atom: true,
  selectable: true,
  parseHTML() {
    return [{ tag: 'div[data-type="page-break"]' }]
  },
  renderHTML({ HTMLAttributes }) {
    return ['div', mergeAttributes(HTMLAttributes, { 'data-type': 'page-break', class: 'sumi-page-break' })]
  },
})

/**
 * Everything a Sumi document can hold: rich Markdown plus Word basics. The editor, the Markdown
 * reader and writer, and the tests all build their schema from this one list.
 */
export const documentExtensions: Extensions = [
  StarterKit.configure({
    heading: { levels: [1, 2, 3] },
    // Clicks place the caret; ⌘-click opens a link (see DocEditor).
    link: { openOnClick: false, defaultProtocol: 'https' },
    // It appends a paragraph on the first transaction, which would mark an untouched document as edited.
    trailingNode: false,
  }),
  TableKit.configure({ table: { resizable: false } }),
  // Inline, like Word and Markdown: an image sits in a paragraph, which carries its alignment.
  Image.configure({ inline: true, allowBase64: true }),
  TextAlign.configure({ types: ['heading', 'paragraph'] }),
  TaskList,
  TaskItem.configure({ nested: true }),
  PageBreak,
]
