/**
 * Reads .docx with mammoth (Word → semantic HTML, images inline as data URLs), then the editor
 * parses that HTML. mammoth has no output for alignment or page breaks, so the importer leaves
 * marker text for them in mammoth's document model and turns the markers into HTML afterwards.
 */

import mammoth from 'mammoth'
import type { ImportFindingInput } from '@shared/fidelity'
import type { PageFormat } from '../page'
import { checklistMarks, docxStyleNames } from './docx-styles'
import { inspectDocx } from './docx-package'

export interface DocxImport {
  readonly html: string
  readonly findings: readonly ImportFindingInput[]
}

/** The parts of mammoth's document model this importer reads or rewrites. */
interface MammothElement {
  readonly type: string
  readonly children?: readonly MammothElement[]
  readonly alignment?: string | null
  readonly breakType?: string
  readonly value?: string
}

// Private-use characters with a prefix no document text contains.
const markerOpen = 'zendo:'
const markerClose = ''
const pageBreakMarker = `${markerOpen}page-break${markerClose}`

const styleMap = [
  // mammoth leaves underline out by default because the web confuses it with links.
  'u => u',
  "p[style-name='Title'] => h1:fresh",
  "p[style-name='Subtitle'] => p:fresh",
  "p[style-name='Intense Quote'] => blockquote > p:fresh",
  `p[style-name='${docxStyleNames.Quote}'] => blockquote > p:fresh`,
  `p[style-name='${docxStyleNames.Code}'] => pre:separator('\\n')`,
  `r[style-name='${docxStyleNames.InlineCode}'] => code`,
  `p[style-name='${docxStyleNames.HorizontalLine}'] => hr:fresh`,
  `p[style-name='${docxStyleNames.ListContinue}'] => p:fresh`,
  `p[style-name='${docxStyleNames.TableText}'] => p:fresh`,
  `p[style-name='${docxStyleNames.Checklist}'] => ul.checklist > li:fresh`,
]

const alignments: Readonly<Record<string, string>> = {
  center: 'center',
  right: 'right',
  end: 'right',
  both: 'justify',
  distribute: 'justify',
}

function isPageBreak(element: MammothElement): boolean {
  return element.type === 'break' && element.breakType === 'page'
}

function hasContent(children: readonly MammothElement[]): boolean {
  return children.some(child => child.type !== 'run' || (child.children ?? []).length > 0)
}

/**
 * Splits a paragraph at its page breaks, since the editor's page break is a block of its own,
 * and puts an alignment marker at the start of each part.
 */
function rewriteParagraph(paragraph: MammothElement): MammothElement[] {
  const segments: Array<MammothElement[] | 'page-break'> = [[]]
  const current = (): MammothElement[] => {
    const last = segments.at(-1)
    if (last === undefined || last === 'page-break') throw new Error('A paragraph segment is missing.')
    return last
  }
  for (const child of paragraph.children ?? []) {
    if (child.type !== 'run' || !(child.children ?? []).some(isPageBreak)) {
      current().push(child)
      continue
    }
    let run: MammothElement[] = []
    for (const grandchild of child.children ?? []) {
      if (!isPageBreak(grandchild)) {
        run.push(grandchild)
        continue
      }
      current().push({ ...child, children: run })
      segments.push('page-break', [])
      run = []
    }
    current().push({ ...child, children: run })
  }

  const alignment = alignments[paragraph.alignment ?? '']
  const alignmentMarker: MammothElement[] = alignment === undefined ? [] : [{ type: 'text', value: `${markerOpen}${alignment}${markerClose}` }]
  const split = segments.length > 1
  return segments.flatMap((segment): MammothElement[] => {
    if (segment === 'page-break') return [{ type: 'paragraph', children: [{ type: 'text', value: pageBreakMarker }] }]
    // A page break at the start or end of a paragraph leaves an empty part that Word never showed.
    if (split && !hasContent(segment)) return []
    return [{ ...paragraph, children: [...alignmentMarker, ...segment] }]
  })
}

function transform(element: MammothElement): MammothElement {
  if (element.children === undefined) return element
  return {
    ...element,
    children: element.children.flatMap(child => child.type === 'paragraph' ? rewriteParagraph(child) : [transform(child)]),
  }
}

interface HtmlResult {
  readonly html: string
  readonly deepHeadings: boolean
  readonly imagesSized: boolean
}

function finishHtml(html: string, imageSizes: readonly { readonly width: number; readonly height: number }[]): HtmlResult {
  let result = html
    .replaceAll(`<p>${pageBreakMarker}</p>`, '<div data-type="page-break"></div>')
    .replace(/<(p|h[1-6])((?:\s[^>]*)?)>zendo:(center|right|justify)/g, '<$1$2 style="text-align: $3">')
    .replace(/zendo:[a-z-]+/g, '')
    .replaceAll('<ul class="checklist">', '<ul data-type="taskList">')
    .replaceAll(`<li>${checklistMarks.open} `, '<li data-type="taskItem" data-checked="false">')
    .replaceAll(`<li>${checklistMarks.done} `, '<li data-type="taskItem" data-checked="true">')
    // Browsers cannot show these formats, and Word could not take them back from the editor either.
    .replace(/<img\b[^>]*src="data:image\/(x-emf|x-wmf|emf|wmf|tiff)[^"]*"[^>]*>/g, '')

  // Sumi has three heading levels; deeper headings become Heading 3.
  const deepHeadings = /<h[4-6]\b/.test(result)
  result = result.replace(/<(\/?)h[4-6]\b/g, '<$1h3')

  // Pictures come out of mammoth in document order without sizes; apply the sizes the package
  // gives, but only when the counts agree, so a size never lands on the wrong picture.
  const images = result.match(/<img\b/g)?.length ?? 0
  const imagesSized = images === imageSizes.length
  if (imagesSized && images > 0) {
    let index = 0
    result = result.replace(/<img\b/g, () => {
      const size = imageSizes[index]
      index += 1
      return size === undefined ? '<img' : `<img width="${size.width}" height="${size.height}"`
    })
  }
  return { html: result, deepHeadings, imagesSized: imagesSized || images === 0 }
}

/** Turns mammoth's warnings into findings: styles it did not recognise and content it skipped. */
function messageFindings(messages: ReadonlyArray<{ readonly type: string; readonly message: string }>): ImportFindingInput[] {
  const paragraphStyles = new Set<string>()
  const runStyles = new Set<string>()
  const ignored = new Set<string>()
  const other = new Set<string>()
  for (const { type, message } of messages) {
    const paragraphStyle = /^Unrecognised paragraph style: '([^']*)'/.exec(message)?.[1]
    const runStyle = /^Unrecognised run style: '([^']*)'/.exec(message)?.[1]
    const element = /^An unrecognised element was ignored: (.+)$/.exec(message)?.[1]
    if (paragraphStyle !== undefined) paragraphStyles.add(paragraphStyle)
    else if (runStyle !== undefined) runStyles.add(runStyle)
    else if (element !== undefined || type === 'error') ignored.add(element ?? message)
    // Image format warnings are already reported from the package's media parts.
    else if (!message.startsWith('Image of type')) other.add(message)
  }
  const findings: ImportFindingInput[] = []
  if (paragraphStyles.size > 0) findings.push({ construct: 'Paragraph styles', severity: 'degraded', location: [...paragraphStyles].join(', '), suggestedAlternative: 'Shown as normal text' })
  if (runStyles.size > 0) findings.push({ construct: 'Character styles', severity: 'degraded', location: [...runStyles].join(', '), suggestedAlternative: 'Shown as plain text' })
  if (ignored.size > 0) findings.push({ construct: 'Unrecognized content', severity: 'dropped', location: [...ignored].join(', ') })
  if (other.size > 0) findings.push({ construct: 'Other formatting', severity: 'degraded', location: [...other].join('; ') })
  return findings
}

export async function readDocx(bytes: Uint8Array, page: PageFormat): Promise<DocxImport> {
  const inspection = inspectDocx(bytes, page)
  const result = await mammoth.convertToHtml(
    { arrayBuffer: bytes.slice().buffer },
    { styleMap, ignoreEmptyParagraphs: false, transformDocument: transform },
  )
  const { html, deepHeadings, imagesSized } = finishHtml(result.value, inspection.imageSizes)
  const findings = [...inspection.findings, ...messageFindings(result.messages)]
  if (deepHeadings) findings.push({ construct: 'Headings below level 3', severity: 'degraded', suggestedAlternative: 'Shown as Heading 3' })
  if (!imagesSized) findings.push({ construct: 'Image sizes', severity: 'degraded', suggestedAlternative: 'Images are shown at their own size, up to the page width' })
  return { html, findings }
}
