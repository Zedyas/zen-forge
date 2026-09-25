/**
 * Markdown through TipTap's own Markdown support (GitHub-flavoured: tables, task lists, images).
 * The writer leaves out what Markdown cannot hold (underline, alignment, page breaks, fonts);
 * saving over a file checks what would change first (doc-documents.ts).
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

/** Links a document may keep: web and mail links, and links relative to the file. Never `javascript:` and the like. */
function safeLink(href: unknown): boolean {
  if (typeof href !== 'string') return false
  const scheme = /^([a-z][a-z0-9+.-]*):/i.exec(href.trim())?.[1]?.toLowerCase()
  return scheme === undefined || scheme === 'http' || scheme === 'https' || scheme === 'mailto'
}

/**
 * Fits parsed Markdown to the schema: TipTap's Markdown reader treats an image alone on a line as
 * a block, but Sumi's images are inline; headings deeper than 3 become Heading 3; unsafe links
 * lose their link.
 */
function normalize(node: JSONContent): JSONContent {
  const level: unknown = node.attrs?.['level']
  const attrs = node.type === 'heading' && typeof level === 'number' && level > 3 ? { ...node.attrs, level: 3 } : node.attrs
  const marks = node.marks?.filter(mark => mark.type !== 'link' || safeLink(mark.attrs?.['href']))
  const content = node.content?.map(child =>
    blockContainers.has(node.type ?? '') && child.type === 'image' ? { type: 'paragraph', content: [child] } : normalize(child))
  return {
    ...node,
    ...(attrs === undefined ? {} : { attrs }),
    ...(marks === undefined ? {} : { marks }),
    ...(content === undefined ? {} : { content }),
  }
}

function places(count: number): string {
  return `${count} ${count === 1 ? 'place' : 'places'}`
}

/** What Markdown can say that a Sumi document cannot, with what opening and saving does to it. */
function markdownFindings(text: string): ImportFindingInput[] {
  let html = 0
  let deepHeadings = 0
  markdown().instance.walkTokens(markdown().instance.lexer(text), (token: MarkdownToken) => {
    if (token.type === 'html') html += 1
    if (token.type === 'heading' && token.depth > 3) deepHeadings += 1
  })
  const findings: ImportFindingInput[] = []
  if (html > 0) findings.push({ construct: 'HTML', location: places(html), suggestedAlternative: 'Shown as the text it contains, and saved as plain Markdown without the HTML.' })
  if (deepHeadings > 0) findings.push({ construct: 'Headings below level 3', location: places(deepHeadings), suggestedAlternative: 'Shown as level 3 headings, and saved that way.' })
  if (/^\[\^[^\]]+\]:/m.test(text)) findings.push({ construct: 'Footnotes', suggestedAlternative: 'Shown as plain text, and saved that way.' })
  return findings
}

/** YAML front matter at the very start of the file, which is kept exactly as it was. */
const frontMatterPattern = /^---\r?\n[\s\S]*?\r?\n---[ \t]*(?:\r?\n|$)/

export function readMarkdown(text: string): MarkdownImport {
  const frontMatter = frontMatterPattern.exec(text)?.[0]
  const body = frontMatter === undefined ? text : text.slice(frontMatter.length)
  const parsed = normalize(markdown().parse(body))
  return {
    content: frontMatter === undefined ? parsed : { ...parsed, attrs: { ...parsed.attrs, frontMatter } },
    findings: markdownFindings(body),
  }
}

/**
 * Characters that must reach the file backslash-escaped. TipTap escapes inline syntax itself but
 * also escapes backslashes, so these are swapped for private-use characters before writing and
 * turned into escapes after.
 */
function placeholder(character: string): string {
  return String.fromCharCode(0xe000 + character.charCodeAt(0))
}

function escapePlaceholders(text: string): string {
  return text.replace(/[-]/g, character => `\\${String.fromCharCode(character.charCodeAt(0) - 0xe000)}`)
}

/** Text that would start a heading, list, rule or setext underline at the start of a line stays text. */
function escapeLineStart(text: string): string {
  return text
    .replace(/^([#+=-])/, (_match, marker: string) => placeholder(marker))
    .replace(/^(\d{1,9})([.)])/, (_match, digits: string, marker: string) => `${digits}${placeholder(marker)}`)
}

function escapeLines(content: readonly JSONContent[], inTable: boolean): JSONContent[] {
  return content.map((child, index) => {
    if (child.type !== 'text' || child.text === undefined) return child
    const code = child.marks?.some(mark => mark.type === 'code') === true
    const lineStart = index === 0 || content[index - 1]?.type === 'hardBreak'
    let text = lineStart && !code ? escapeLineStart(child.text) : child.text
    // A pipe inside a table cell ends the cell unless escaped, even inside code.
    if (inTable) text = text.replaceAll('|', placeholder('|'))
    return { ...child, text }
  })
}

/** Markdown tables always start with a header row; a table without one uses its first row, so no empty row is added. */
function withHeaderRow(table: JSONContent): JSONContent {
  const [first, ...rest] = table.content ?? []
  if (first === undefined || first.content?.some(cell => cell.type === 'tableHeader')) return table
  const header = { ...first, content: first.content?.map(cell => ({ ...cell, type: 'tableHeader' })) }
  return { ...table, content: [header, ...rest] }
}

function spanOf(cell: JSONContent, name: string): number {
  const value: unknown = cell.attrs?.[name]
  return typeof value === 'number' && value > 1 ? value : 1
}

/**
 * Markdown has no merged cells: a merged cell keeps its value in its first column, and the
 * columns and rows it covered get empty cells, so every other value stays in its column.
 */
function withoutMerges(table: JSONContent): JSONContent {
  const covered = new Map<number, number>()
  const rows = (table.content ?? []).map(row => {
    const type = row.content?.[0]?.type ?? 'tableCell'
    const empty = (): JSONContent => ({ type, content: [{ type: 'paragraph' }] })
    const cells: JSONContent[] = []
    const fill = (): void => {
      while ((covered.get(cells.length) ?? 0) > 0) {
        covered.set(cells.length, (covered.get(cells.length) ?? 1) - 1)
        cells.push(empty())
      }
    }
    for (const cell of row.content ?? []) {
      fill()
      const colspan = spanOf(cell, 'colspan')
      const rowspan = spanOf(cell, 'rowspan')
      const start = cells.length
      cells.push({ ...cell, attrs: { ...cell.attrs, colspan: 1, rowspan: 1, colwidth: null } })
      for (let extra = 1; extra < colspan; extra += 1) cells.push(empty())
      if (rowspan > 1) for (let column = start; column < start + colspan; column += 1) covered.set(column, rowspan - 1)
    }
    fill()
    return { ...row, content: cells }
  })
  return { ...table, content: rows }
}

/** The marks Markdown has syntax for; fonts, colours, underline and the rest are left out. */
const markdownMarks = new Set(['bold', 'italic', 'strike', 'code', 'link'])

/**
 * The document as Markdown can hold it: no page breaks, no merged cells, no formatting beyond
 * bold, italic, strikethrough, code and links. `imageSources` maps an image shown from a data URL
 * back to the path the file had for it.
 */
function forMarkdown(node: JSONContent, imageSources: ReadonlyMap<string, string>, inTable = false): JSONContent {
  if (node.type === 'table' && node.content !== undefined) {
    return { ...node, content: withoutMerges(withHeaderRow(node)).content?.map(child => forMarkdown(child, imageSources, true)) }
  }
  const src: unknown = node.attrs?.['src']
  const original = node.type === 'image' && typeof src === 'string' ? imageSources.get(src) : undefined
  const children = node.content?.filter(child => child.type !== 'pageBreak').map(child => forMarkdown(child, imageSources, inTable))
  const textblock = node.type === 'paragraph' || node.type === 'heading' || node.type === 'title'
  return {
    ...node,
    ...(original === undefined ? {} : { attrs: { ...node.attrs, src: original } }),
    ...(node.marks === undefined ? {} : { marks: node.marks.filter(mark => markdownMarks.has(mark.type) && (mark.type !== 'link' || safeLink(mark.attrs?.['href']))) }),
    ...(children === undefined ? {} : { content: textblock ? escapeLines(children, inTable) : children }),
  }
}

export function writeMarkdown(content: JSONContent, imageSources: ReadonlyMap<string, string> = new Map()): string {
  const frontMatter: unknown = content.attrs?.['frontMatter']
  const body = `${escapePlaceholders(markdown().serialize(forMarkdown(content, imageSources)).trim())}\n`
  return typeof frontMatter === 'string' && frontMatter !== '' ? `${frontMatter.endsWith('\n') ? frontMatter : `${frontMatter}\n`}\n${body}` : body
}
