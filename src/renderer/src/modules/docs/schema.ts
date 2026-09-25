import { Extension, mergeAttributes, Node, type Attribute, type Extensions } from '@tiptap/core'
import Image from '@tiptap/extension-image'
import { TaskItem, TaskList } from '@tiptap/extension-list'
import { Subscript } from '@tiptap/extension-subscript'
import { Superscript } from '@tiptap/extension-superscript'
import { Table, TableCell, TableHeader, TableRow } from '@tiptap/extension-table'
import TextAlign from '@tiptap/extension-text-align'
import { BackgroundColor, Color, FontFamily, FontSize, TextStyle } from '@tiptap/extension-text-style'
import StarterKit from '@tiptap/starter-kit'
import { defaultPage, defaultTheme, emptyProperties, singleLineHeight, toPoints } from './theme'

/**
 * The root node. Its attributes are the document's page setup, theme and file properties
 * (theme.ts), and a Markdown file's front matter, kept exactly as it was.
 */
const SumiDocument = Node.create({
  name: 'doc',
  topNode: true,
  content: 'block+',
  addAttributes() {
    return {
      page: { default: defaultPage, rendered: false },
      theme: { default: defaultTheme, rendered: false },
      properties: { default: emptyProperties, rendered: false },
      frontMatter: { default: null, rendered: false },
    }
  },
  renderMarkdown: (node, helpers) => helpers.renderChildren(node.content ?? [], '\n\n'),
})

/** Word's Title style. Markdown has no title, so it is written as a first-level heading. */
export const Title = Node.create({
  name: 'title',
  group: 'block',
  content: 'inline*',
  defining: true,
  parseHTML() {
    return [{ tag: 'p[data-style="title"]', priority: 60 }]
  },
  renderHTML({ HTMLAttributes }) {
    return ['p', mergeAttributes(HTMLAttributes, { 'data-style': 'title' }), 0]
  },
  renderMarkdown: (node, helpers) => `# ${helpers.renderChildren(node.content ?? [])}`,
})

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

/** A length in points, stored as a number and shown as one CSS property. */
function pointAttribute(name: string, property: 'margin-top' | 'margin-bottom' | 'margin-left' | 'text-indent'): Attribute {
  return {
    default: null,
    parseHTML: element => toPoints(element.style.getPropertyValue(property)) ?? null,
    renderHTML: attributes => {
      const value: unknown = attributes[name]
      return typeof value === 'number' ? { style: `${property}: ${value}pt` } : {}
    },
  }
}

/**
 * Direct paragraph formatting on top of the style: line spacing (a multiple of single spacing),
 * space before and after, left indent and first-line indent (negative for a hanging indent).
 */
const ParagraphFormat = Extension.create({
  name: 'paragraphFormat',
  addGlobalAttributes() {
    return [{
      types: ['paragraph', 'heading', 'title'],
      attributes: {
        lineHeight: {
          default: null,
          parseHTML: element => {
            const value = Number.parseFloat(element.style.lineHeight)
            return Number.isFinite(value) ? Math.round((value / singleLineHeight) * 100) / 100 : null
          },
          renderHTML: attributes => {
            const value: unknown = attributes['lineHeight']
            return typeof value === 'number' ? { style: `line-height: ${value * singleLineHeight}` } : {}
          },
        },
        spaceBefore: pointAttribute('spaceBefore', 'margin-top'),
        spaceAfter: pointAttribute('spaceAfter', 'margin-bottom'),
        indentLeft: pointAttribute('indentLeft', 'margin-left'),
        indentFirstLine: pointAttribute('indentFirstLine', 'text-indent'),
      },
    }]
  },
})

/** Cell shading, as Word's table cells have it. */
const cellShading: Record<string, Attribute> = {
  backgroundColor: {
    default: null,
    parseHTML: element => element.getAttribute('data-background'),
    renderHTML: attributes => {
      const value: unknown = attributes['backgroundColor']
      return typeof value === 'string' ? { 'data-background': value, style: `background-color: ${value}` } : {}
    },
  },
}

const ShadedTableCell = TableCell.extend({
  addAttributes() {
    return { ...this.parent?.(), ...cellShading }
  },
})

const ShadedTableHeader = TableHeader.extend({
  addAttributes() {
    return { ...this.parent?.(), ...cellShading }
  },
})

/**
 * Everything a Sumi document can hold: rich Markdown plus Word's everyday formatting. The editor,
 * the Markdown reader and writer, and the tests all build their schema from this one list.
 */
export const documentExtensions: Extensions = [
  SumiDocument,
  StarterKit.configure({
    document: false,
    heading: { levels: [1, 2, 3] },
    // Clicks place the caret; ⌘-click opens a link (see DocEditor).
    link: { openOnClick: false, defaultProtocol: 'https' },
    // It appends a paragraph on the first transaction, which would mark an untouched document as edited.
    trailingNode: false,
  }),
  Title,
  ParagraphFormat,
  TextAlign.configure({ types: ['heading', 'paragraph', 'title'] }),
  TextStyle,
  FontFamily,
  FontSize,
  Color,
  BackgroundColor,
  Superscript,
  Subscript,
  Table.configure({ resizable: true, cellMinWidth: 36 }),
  TableRow,
  ShadedTableHeader,
  ShadedTableCell,
  // Inline, like Word and Markdown: an image sits in a paragraph, which carries its alignment.
  Image.configure({
    inline: true,
    allowBase64: true,
    resize: { enabled: true, alwaysPreserveAspectRatio: true, directions: ['top-left', 'top-right', 'bottom-left', 'bottom-right'], minWidth: 24, minHeight: 24 },
  }),
  TaskList,
  TaskItem.configure({ nested: true }),
  PageBreak,
]
