import { Extension, mergeAttributes, Node, type Attribute, type Extensions } from '@tiptap/core'
import CodeBlock from '@tiptap/extension-code-block'
import Image from '@tiptap/extension-image'
import { TaskItem, TaskList } from '@tiptap/extension-list'
import { Subscript } from '@tiptap/extension-subscript'
import { Superscript } from '@tiptap/extension-superscript'
import { Table, TableCell, TableHeader, TableRow } from '@tiptap/extension-table'
import TextAlign from '@tiptap/extension-text-align'
import { BackgroundColor, Color, FontFamily, FontSize, TextStyle } from '@tiptap/extension-text-style'
import StarterKit from '@tiptap/starter-kit'
import { noHeadersOrFooters } from './header-footer'
import { contentsEntries, contentsLevels } from './text-stats'
import { defaultPage, defaultTheme, emptyProperties, singleLineHeight, toPoints } from './theme'

/**
 * The root node. Its attributes are the document's page setup, theme and file properties
 * (theme.ts), its headers and footers (header-footer.ts), and a Markdown file's front matter,
 * kept exactly as it was.
 */
const SumiDocument = Node.create({
  name: 'doc',
  topNode: true,
  // A table of contents only at the top level: inside a list or quotation, Word's field would end the list or quotation.
  content: '(block | tableOfContents)+',
  addAttributes() {
    return {
      page: { default: defaultPage, rendered: false },
      theme: { default: defaultTheme, rendered: false },
      properties: { default: emptyProperties, rendered: false },
      headerFooter: { default: noHeadersOrFooters, rendered: false },
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

/**
 * A table of contents, as Word's TOC field: the headings from Heading 1 down to `levels`. It holds
 * no text of its own; the editor lists the headings as they are now, and clicking one goes there.
 * Word gets a TOC field with the entries as its last result (docx-export.ts), print fills them in
 * (print.ts), and Markdown gets a list of links (markdown.ts).
 */
export const TableOfContents = Node.create({
  name: 'tableOfContents',
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      levels: {
        default: 3,
        parseHTML: element => contentsLevels(Number(element.getAttribute('data-levels'))),
        renderHTML: attributes => ({ 'data-levels': contentsLevels(attributes['levels']) }),
      },
    }
  },
  parseHTML() {
    return [{ tag: 'div[data-type="table-of-contents"]' }]
  },
  renderHTML({ HTMLAttributes }) {
    return ['div', mergeAttributes(HTMLAttributes, { 'data-type': 'table-of-contents', class: 'sumi-toc' })]
  },
  addNodeView() {
    return ({ editor, node }) => {
      const dom = document.createElement('div')
      dom.className = 'sumi-toc'
      dom.dataset['type'] = 'table-of-contents'
      dom.contentEditable = 'false'
      let levels = contentsLevels(node.attrs['levels'])
      let shown: string | undefined
      // Redrawn when a heading's text or level changes; positions are looked up again on click.
      const render = (): void => {
        const entries = contentsEntries(editor.state.doc, levels)
        const key = JSON.stringify([levels, entries.map(entry => [entry.level, entry.text])])
        if (key === shown) return
        shown = key
        if (entries.length === 0) {
          const empty = document.createElement('p')
          empty.className = 'sumi-toc-empty'
          empty.textContent = `Table of contents: headings 1 to ${levels} are listed here.`
          dom.replaceChildren(empty)
          return
        }
        dom.replaceChildren(...entries.map((entry, index) => {
          const button = document.createElement('button')
          button.type = 'button'
          button.className = 'sumi-toc-entry'
          button.dataset['level'] = String(entry.level)
          button.textContent = entry.text === '' ? 'Untitled heading' : entry.text
          button.addEventListener('click', () => {
            const target = contentsEntries(editor.state.doc, levels)[index]
            if (target !== undefined) editor.chain().focus().setTextSelection(target.position).scrollIntoView().run()
          })
          return button
        }))
      }
      render()
      editor.on('update', render)
      return {
        dom,
        update: updated => {
          if (updated.type.name !== 'tableOfContents') return false
          levels = contentsLevels(updated.attrs['levels'])
          render()
          return true
        },
        // Clicks on an entry are the entry's; elsewhere they select the table, so it can be deleted.
        stopEvent: event => event.target instanceof Element && event.target.closest('.sumi-toc-entry') !== null,
        ignoreMutation: () => true,
        destroy: () => {
          editor.off('update', render)
        },
      }
    }
  },
})

/**
 * A code block whose Markdown fence is longer than any run of backticks inside it, so code that
 * shows a fenced block is written back intact.
 */
const FencedCodeBlock = CodeBlock.extend({
  renderMarkdown: (node, helpers) => {
    const text = helpers.renderChildren(node.content ?? [])
    const longest = Math.max(2, ...[...text.matchAll(/`+/g)].map(match => match[0].length))
    const fence = '`'.repeat(longest + 1)
    const language: unknown = node.attrs?.['language']
    return `${fence}${typeof language === 'string' ? language : ''}\n${text}\n${fence}`
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
    codeBlock: false,
    heading: { levels: [1, 2, 3, 4, 5, 6] },
    // Clicks place the caret; ⌘-click opens a link (see DocEditor).
    link: { openOnClick: false, defaultProtocol: 'https' },
    // It appends a paragraph on the first transaction, which would mark an untouched document as edited.
    trailingNode: false,
  }),
  FencedCodeBlock,
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
  TableOfContents,
]

/** A header or footer: paragraphs only. */
const HeaderFooterDocument = Node.create({
  name: 'doc',
  topNode: true,
  content: 'paragraph+',
})

/** The page number or the page count, as Word's PAGE and NUMPAGES fields; each page shows its own. */
export const PageField = Node.create({
  name: 'pageField',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      kind: {
        default: 'page',
        parseHTML: element => element.getAttribute('data-field') === 'total' ? 'total' : 'page',
        renderHTML: attributes => ({ 'data-field': attributes['kind'] === 'total' ? 'total' : 'page' }),
      },
    }
  },
  parseHTML() {
    return [{ tag: 'span[data-field]' }]
  },
  renderHTML({ node, HTMLAttributes }) {
    return ['span', mergeAttributes(HTMLAttributes, { class: 'sumi-field' }), node.attrs['kind'] === 'total' ? 'pages' : 'page']
  },
})

/** What a header or footer can hold (header-footer.ts). Its small editor and the .docx reader and writer use this list. */
export const headerFooterExtensions: Extensions = [
  HeaderFooterDocument,
  StarterKit.configure({
    document: false,
    heading: false,
    blockquote: false,
    bulletList: false,
    orderedList: false,
    listItem: false,
    listKeymap: false,
    codeBlock: false,
    code: false,
    horizontalRule: false,
    strike: false,
    link: false,
    dropcursor: false,
    gapcursor: false,
    trailingNode: false,
  }),
  TextAlign.configure({ types: ['paragraph'] }),
  TextStyle,
  FontFamily,
  FontSize,
  Color,
  BackgroundColor,
  PageField,
]
