import type { Editor, JSONContent } from '@tiptap/core'
import { DOMSerializer } from '@tiptap/pm/model'
import { toast } from 'sonner'
import { openFiles } from '../../app/document-actions'
import { useDocumentsStore } from '../../app/documents-store'
import { fileService } from '../../services/file/IpcFileService'
import { openPrintDialog, printToPdf, type Printout } from '../../services/print/print'
import { documentHeadersAndFooters, type HeaderFooterContent } from './header-footer'
import { hexColor } from './io/docx-styles'
import { documentPage, documentTheme, themeVariables, toPoints, type BlockStyle } from './theme'

const millimetresPerTwip = 25.4 / 1440

/** Text as a CSS string; a line end is CSS's `\A`. */
function cssString(text: string): string {
  return `"${text.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\r?\n/g, '\\A ')}"`
}

/** One paragraph of a header or footer as a CSS `content` value, with page numbers as counters. */
function boxContent(paragraph: JSONContent): string {
  const parts = (paragraph.content ?? []).map(node => {
    if (node.type === 'pageField') return node.attrs?.['kind'] === 'total' ? 'counter(pages)' : 'counter(page)'
    if (node.type === 'hardBreak') return '"\\A"'
    return cssString(node.text ?? '')
  })
  return parts.length === 0 ? '""' : parts.join(' ')
}

/** A margin box holds text in one format: the paragraph style's, with its first run's formatting over it. */
function boxFont(paragraph: JSONContent | undefined, base: BlockStyle): string {
  const first = paragraph?.content?.find(node => node.type === 'text' || node.type === 'pageField')
  const has = (type: string): boolean => first?.marks?.some(mark => mark.type === type) === true
  const style = first?.marks?.find(mark => mark.type === 'textStyle')?.attrs ?? {}
  const family: unknown = style['fontFamily']
  const color: unknown = style['color']
  return [
    `font-family: ${cssString(typeof family === 'string' && family !== '' ? family : base.fontFamily)}, sans-serif;`,
    `font-size: ${toPoints(style['fontSize']) ?? base.fontSize}pt;`,
    `color: ${(typeof color === 'string' ? hexColor(color) : undefined) ?? base.color};`,
    has('bold') ? 'font-weight: 700;' : '',
    has('italic') ? 'font-style: italic;' : '',
    has('underline') ? 'text-decoration: underline;' : '',
  ].join(' ')
}

const boxSides = ['left', 'center', 'right'] as const

/**
 * A header or footer as the @page margin boxes along one edge: each paragraph goes in the box for
 * its alignment (justified ones on the left), and page numbers are the page counters, so every
 * page shows its own. A side with nothing is `none`, which clears what the general page rule set.
 */
function marginBoxes(edge: 'top' | 'bottom', content: HeaderFooterContent, base: BlockStyle): string {
  return boxSides.map(side => {
    const paragraphs = content.filter(paragraph => {
      const align: unknown = paragraph.attrs?.['textAlign']
      return side === 'left' ? align !== 'center' && align !== 'right' : align === side
    })
    const value = paragraphs.length === 0 ? 'none' : paragraphs.map(boxContent).join(' "\\A" ')
    return `@${edge}-${side} { content: ${value}; white-space: pre-wrap; text-align: ${side}; vertical-align: middle; line-height: 1.15; ${boxFont(paragraphs[0], base)} }`
  }).join('\n')
}

/**
 * A picture that is not stored in the document, or whose data cannot be shown, prints as its
 * description instead, so one broken picture never stops the printout.
 */
async function replaceBrokenImages(root: HTMLElement): Promise<void> {
  await Promise.all(Array.from(root.querySelectorAll('img'), async image => {
    const stored = (image.getAttribute('src') ?? '').startsWith('data:')
    if (stored && await image.decode().then(() => true, () => false)) return
    const description = document.createElement('span')
    description.className = 'sumi-print-missing'
    description.textContent = image.alt === '' ? '[Picture]' : `[${image.alt}]`
    image.replaceWith(description)
  }))
}

/**
 * The document as it prints: the editor's content serialized by the schema, styled by the same
 * `.sumi-prose` rules and theme variables as the page on screen (docs.css), on the document's own
 * paper. Word margins can differ on each side, and the header and footer sit in the top and
 * bottom margins, so the printout carries its own @page rules for both; the print service's rule
 * sets the paper size.
 */
export async function docPrintout(editor: Editor): Promise<Printout> {
  const doc = editor.state.doc
  const page = documentPage(doc.attrs['page'])
  const theme = documentTheme(doc.attrs['theme'])
  const headerFooter = documentHeadersAndFooters(doc.attrs['headerFooter'])
  const margin = (twips: number): string => `${(twips * millimetresPerTwip).toFixed(2)}mm`
  const boxes = (header: HeaderFooterContent, footer: HeaderFooterContent): string => `${marginBoxes('top', header, theme.normal)}\n${marginBoxes('bottom', footer, theme.normal)}`
  const rule = document.createElement('style')
  rule.textContent = [
    `@page { margin: ${margin(page.marginTop)} ${margin(page.marginRight)} ${margin(page.marginBottom)} ${margin(page.marginLeft)};\n${boxes(headerFooter.header, headerFooter.footer)} }`,
    ...(headerFooter.differentFirstPage ? [`@page :first {\n${boxes(headerFooter.firstHeader, headerFooter.firstFooter)} }`] : []),
  ].join('\n')

  const prose = document.createElement('div')
  prose.className = 'sumi-prose sumi-print'
  for (const [name, value] of themeVariables(theme)) prose.style.setProperty(name, value)
  prose.append(DOMSerializer.fromSchema(editor.schema).serializeFragment(doc.content))
  await replaceBrokenImages(prose)

  const content = document.createElement('div')
  content.append(rule, prose)
  return { content, page: { width: page.width * millimetresPerTwip, height: page.height * millimetresPerTwip, margin: 0 } }
}

export async function printDoc(editor: Editor): Promise<void> {
  await openPrintDialog(() => docPrintout(editor))
}

/** Writes the document, as it prints, to a PDF chosen by the user, and offers to open it in Hanko. */
export async function exportDocPdf(documentId: string, editor: Editor): Promise<void> {
  const document = useDocumentsStore.getState().documents.find(candidate => candidate.id === documentId)
  if (document === undefined) return
  const path = await fileService.chooseSavePath({ defaultName: `${document.name}.pdf`, extensions: ['pdf'] })
  if (path === undefined) return
  const bytes = await printToPdf(() => docPrintout(editor))
  if (bytes === undefined) return
  await fileService.write(path, bytes)
  const file = await fileService.describe(path)
  toast.success('Exported as PDF', {
    description: file.name,
    action: { label: 'Open', onClick: () => void openFiles([file]) },
  })
}
