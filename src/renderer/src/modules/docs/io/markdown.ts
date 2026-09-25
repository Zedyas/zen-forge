/**
 * Markdown through TipTap's own Markdown support (GitHub-flavoured: tables, task lists, images).
 * The writer leaves out what Markdown cannot hold (underline, alignment, page breaks).
 */

import type { JSONContent } from '@tiptap/core'
import { MarkdownManager } from '@tiptap/markdown'
import type { ImportFindingInput } from '@shared/fidelity'
import { documentExtensions } from '../schema'

export interface MarkdownImport {
  readonly content: JSONContent
  readonly findings: readonly ImportFindingInput[]
}

let manager: MarkdownManager | undefined

function markdown(): MarkdownManager {
  manager ??= new MarkdownManager({ extensions: documentExtensions })
  return manager
}

type MarkdownToken = ReturnType<MarkdownManager['instance']['lexer']>[number]

/** Containers whose children must be blocks; an image on its own line becomes a paragraph there. */
const blockContainers = new Set(['doc', 'blockquote', 'listItem', 'taskItem', 'tableCell', 'tableHeader'])

/**
 * Fits parsed Markdown to the schema: TipTap's Markdown reader treats an image alone on a line as
 * a block, but Sumi's images are inline; headings deeper than 3 become Heading 3.
 */
function normalize(node: JSONContent): JSONContent {
  const level: unknown = node.attrs?.['level']
  const attrs = node.type === 'heading' && typeof level === 'number' && level > 3 ? { ...node.attrs, level: 3 } : node.attrs
  const content = node.content?.map(child =>
    blockContainers.has(node.type ?? '') && child.type === 'image' ? { type: 'paragraph', content: [child] } : normalize(child))
  return { ...node, ...(attrs === undefined ? {} : { attrs }), ...(content === undefined ? {} : { content }) }
}

function places(count: number): string {
  return `${count} ${count === 1 ? 'place' : 'places'}`
}

/** What Markdown can say that a Sumi document cannot, with what opening and saving does to it. */
function markdownFindings(text: string): ImportFindingInput[] {
  let html = 0
  let deepHeadings = 0
  let alignedTables = 0
  markdown().instance.walkTokens(markdown().instance.lexer(text), (token: MarkdownToken) => {
    if (token.type === 'html') html += 1
    if (token.type === 'heading' && token.depth > 3) deepHeadings += 1
    if (token.type === 'table' && token.align.some((align: unknown) => align !== null)) alignedTables += 1
  })
  const findings: ImportFindingInput[] = []
  if (html > 0) findings.push({ construct: 'HTML', location: places(html), suggestedAlternative: 'Shown as the text it contains; saved as plain Markdown without the HTML.' })
  if (deepHeadings > 0) findings.push({ construct: 'Headings below level 3', location: places(deepHeadings), suggestedAlternative: 'Shown and saved as level 3 headings.' })
  if (alignedTables > 0) findings.push({ construct: 'Table column alignment', location: places(alignedTables), suggestedAlternative: 'Columns are shown and saved aligned left.' })
  if (/^---\r?\n[\s\S]*?\r?\n---\r?\n/.test(text)) {
    findings.push({ construct: 'Front matter', suggestedAlternative: 'Shown and saved as ordinary text, so tools that read front matter will not find it.' })
  }
  if (/^\[\^[^\]]+\]:/m.test(text)) findings.push({ construct: 'Footnotes', suggestedAlternative: 'Shown and saved as plain text.' })
  return findings
}

export function readMarkdown(text: string): MarkdownImport {
  return { content: normalize(markdown().parse(text)), findings: markdownFindings(text) }
}

/** Markdown tables always start with a header row; a table without one uses its first row, so no empty row is added. */
function withHeaderRow(table: JSONContent): JSONContent {
  const [first, ...rest] = table.content ?? []
  if (first === undefined || first.content?.some(cell => cell.type === 'tableHeader')) return table
  const header = { ...first, content: first.content?.map(cell => ({ ...cell, type: 'tableHeader' })) }
  return { ...table, content: [header, ...rest] }
}

/** The marks Markdown has syntax for; fonts, colours, underline and the rest are left out. */
const markdownMarks = new Set(['bold', 'italic', 'strike', 'code', 'link'])

/**
 * The document as Markdown can hold it: no page breaks, no formatting beyond bold, italic,
 * strikethrough, code and links. `imageSources` maps an image shown from a data URL back to the
 * path the file had for it.
 */
function forMarkdown(node: JSONContent, imageSources: ReadonlyMap<string, string>): JSONContent {
  if (node.type === 'table' && node.content !== undefined) {
    return { ...node, content: withHeaderRow(node).content?.map(child => forMarkdown(child, imageSources)) }
  }
  const src: unknown = node.attrs?.['src']
  const original = node.type === 'image' && typeof src === 'string' ? imageSources.get(src) : undefined
  return {
    ...node,
    ...(original === undefined ? {} : { attrs: { ...node.attrs, src: original } }),
    ...(node.marks === undefined ? {} : { marks: node.marks.filter(mark => markdownMarks.has(mark.type)) }),
    ...(node.content === undefined ? {} : { content: node.content.filter(child => child.type !== 'pageBreak').map(child => forMarkdown(child, imageSources)) }),
  }
}

export function writeMarkdown(content: JSONContent, imageSources: ReadonlyMap<string, string> = new Map()): string {
  return `${markdown().serialize(forMarkdown(content, imageSources)).trim()}\n`
}
