/**
 * Headers and footers, as plain data on the `doc` node (its `headerFooter` attribute), like the
 * page setup. Each part is a list of paragraphs in the header schema (schema.ts,
 * headerFooterExtensions): text with bold, italic, underline, font, size and colour, alignment,
 * and page number fields. Word's odd and even pages and per-section headers are not kept.
 */

import type { JSONContent } from '@tiptap/core'

export type HeaderFooterContent = readonly JSONContent[]

export interface HeadersAndFooters {
  readonly header: HeaderFooterContent
  readonly footer: HeaderFooterContent
  /** Shown on the first page instead of the others while `differentFirstPage` is on; kept while it is off, as Word does. */
  readonly firstHeader: HeaderFooterContent
  readonly firstFooter: HeaderFooterContent
  readonly differentFirstPage: boolean
}

export type HeaderFooterPart = 'header' | 'footer' | 'firstHeader' | 'firstFooter'

export const headerFooterParts: readonly HeaderFooterPart[] = ['header', 'footer', 'firstHeader', 'firstFooter']

export const partLabels: Readonly<Record<HeaderFooterPart, string>> = {
  header: 'Header',
  footer: 'Footer',
  firstHeader: 'First page header',
  firstFooter: 'First page footer',
}

export const noHeadersOrFooters: HeadersAndFooters = { header: [], footer: [], firstHeader: [], firstFooter: [], differentFirstPage: false }

export type PageField = 'page' | 'total'

/** Word's field for each page number field; NUMPAGES counts every page, SECTIONPAGES the section's, which is the same with one section. */
export const pageFieldCodes: Readonly<Record<string, PageField>> = { PAGE: 'page', NUMPAGES: 'total', SECTIONPAGES: 'total' }

function isRecord(value: unknown): value is Readonly<Record<string, unknown>> {
  return typeof value === 'object' && value !== null
}

function isContent(value: unknown): value is HeaderFooterContent {
  return Array.isArray(value) && value.every(isRecord)
}

/** The headers and footers from the `doc` node's attributes, or none. */
export function documentHeadersAndFooters(value: unknown): HeadersAndFooters {
  if (!isRecord(value)) return noHeadersOrFooters
  const part = (name: HeaderFooterPart): HeaderFooterContent => {
    const content = value[name]
    return isContent(content) ? content : []
  }
  return {
    header: part('header'),
    footer: part('footer'),
    firstHeader: part('firstHeader'),
    firstFooter: part('firstFooter'),
    differentFirstPage: value['differentFirstPage'] === true,
  }
}

/** Paragraphs with nothing in them are no header at all, so an emptied header is not saved. */
export function trimmedPart(content: HeaderFooterContent): HeaderFooterContent {
  return content.every(paragraph => (paragraph.content ?? []).length === 0) ? [] : content
}

function isPageNumber(node: JSONContent): boolean {
  return node.type === 'pageField' && node.attrs?.['kind'] !== 'total'
}

export function hasPageNumber(content: HeaderFooterContent): boolean {
  return content.some(paragraph => (paragraph.content ?? []).some(isPageNumber))
}

/**
 * The footer with a centred page number added, or with its page numbers taken out; a paragraph
 * that held only a page number goes with it.
 */
export function withPageNumber(content: HeaderFooterContent, on: boolean): HeaderFooterContent {
  if (on) return [...content, { type: 'paragraph', attrs: { textAlign: 'center' }, content: [{ type: 'pageField', attrs: { kind: 'page' } }] }]
  return content.flatMap(paragraph => {
    const children = paragraph.content ?? []
    const kept = children.filter(node => !isPageNumber(node))
    if (kept.length === children.length) return [paragraph]
    return kept.length === 0 ? [] : [{ ...paragraph, content: kept }]
  })
}
