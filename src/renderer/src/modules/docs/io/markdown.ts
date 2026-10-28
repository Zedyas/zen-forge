/**
 * Markdown through TipTap's own Markdown support (GitHub-flavoured: tables, task lists, images).
 * The writer leaves out what Markdown cannot hold (underline, alignment, page breaks, fonts,
 * headers and footers) and writes a table of contents as a list of links to the headings;
 * saving over a file checks what would change first (doc-documents.ts).
 */

import type { JSONContent } from '@tiptap/core'
import { MarkdownManager } from '@tiptap/markdown'
import type { ImportFindingInput } from '@shared/fidelity'
import { documentHeadersAndFooters, headerFooterParts } from '../header-footer'
import { documentExtensions } from '../schema'
import { contentsLevels } from '../text-stats'

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
 * a block, but Sumi's images are inline; unsafe links lose their link.
 */
function normalize(node: JSONContent): JSONContent {
  const marks = node.marks?.filter(mark => mark.type !== 'link' || safeLink(mark.attrs?.['href']))
  const content = node.content?.map(child =>
    blockContainers.has(node.type ?? '') && child.type === 'image' ? { type: 'paragraph', content: [child] } : normalize(child))
  return {
    ...node,
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
  markdown().instance.walkTokens(markdown().instance.lexer(text), (token: MarkdownToken) => {
    if (token.type === 'html') html += 1
  })
  const findings: ImportFindingInput[] = []
  if (html > 0) findings.push({ construct: 'HTML', location: places(html), suggestedAlternative: 'Shown as the text it contains, and saved as plain Markdown without the HTML.' })
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

/** What saving as Markdown leaves out that is easy to miss, said in the question before saving over a file. */
export function markdownLosses(content: JSONContent): string[] {
  const headerFooter = documentHeadersAndFooters(content.attrs?.['headerFooter'])
  const contents = (node: JSONContent): boolean => node.type === 'tableOfContents' || (node.content ?? []).some(contents)
  return [
    ...headerFooterParts.some(part => headerFooter[part].length > 0) ? ['the header and footer are left out'] : [],
    ...contents(content) ? ['the table of contents becomes a list of links that no longer updates'] : [],
  ]
}

interface Heading {
  readonly level: number
  readonly text: string
  readonly slug: string
}

/**
 * The document's headings with the anchors GitHub gives them (github-slugger): lowercase,
 * punctuation removed, spaces as hyphens, and -1, -2 after a repeat, skipping any anchor already
 * taken. The title is written as a heading too, so it takes an anchor; it is level 0, which no
 * table of contents lists.
 */
function headings(content: JSONContent): Heading[] {
  const found: Heading[] = []
  const repeats = new Map<string, number>()
  const visit = (node: JSONContent): void => {
    if (node.type === 'heading' || node.type === 'title') {
      const text = (node.content ?? []).map(child => child.text ?? '').join('').trim()
      const level: unknown = node.attrs?.['level']
      const base = text.toLowerCase().replace(/[^\p{L}\p{M}\p{N}\s_-]/gu, '').replace(/\s/g, '-')
      let slug = base
      while (repeats.has(slug)) {
        const count = (repeats.get(base) ?? 0) + 1
        repeats.set(base, count)
        slug = `${base}-${count}`
      }
      repeats.set(slug, 0)
      if (text !== '') found.push({ level: node.type === 'title' ? 0 : typeof level === 'number' ? level : 1, text, slug })
    }
    node.content?.forEach(visit)
  }
  visit(content)
  return found
}

interface ContentsItem {
  readonly entry: Heading
  readonly children: ContentsItem[]
}

function bulletList(items: readonly ContentsItem[]): JSONContent[] {
  if (items.length === 0) return []
  return [{
    type: 'bulletList',
    content: items.map(({ entry, children }) => {
      const link = { type: 'text', text: entry.text, marks: [{ type: 'link', attrs: { href: `#${entry.slug}` } }] }
      return { type: 'listItem', content: [{ type: 'paragraph', content: [link] }, ...bulletList(children)] }
    }),
  }]
}

/** A table of contents as a bulleted list of links to the headings; each heading nests under the last one above its level. */
function contentsList(entries: readonly Heading[]): JSONContent[] {
  const top: ContentsItem[] = []
  const open: ContentsItem[] = []
  for (const entry of entries) {
    while ((open.at(-1)?.entry.level ?? 0) >= entry.level) open.pop()
    const item: ContentsItem = { entry, children: [] }
    const parent = open.at(-1)
    if (parent === undefined) top.push(item)
    else parent.children.push(item)
    open.push(item)
  }
  return bulletList(top)
}

/** The document with each table of contents written out as a list of links; it no longer updates. */
function withContentsLists(node: JSONContent, entries: readonly Heading[]): JSONContent {
  if (node.content === undefined) return node
  return {
    ...node,
    content: node.content.flatMap(child => child.type === 'tableOfContents'
      ? contentsList(entries.filter(entry => entry.level >= 1 && entry.level <= contentsLevels(child.attrs?.['levels'])))
      : [withContentsLists(child, entries)]),
  }
}

export function writeMarkdown(content: JSONContent, imageSources: ReadonlyMap<string, string> = new Map()): string {
  const frontMatter: unknown = content.attrs?.['frontMatter']
  const written = forMarkdown(withContentsLists(content, headings(content)), imageSources)
  const body = `${escapePlaceholders(markdown().serialize(written).trim())}\n`
  return typeof frontMatter === 'string' && frontMatter !== '' ? `${frontMatter.endsWith('\n') ? frontMatter : `${frontMatter}\n`}\n${body}` : body
}
