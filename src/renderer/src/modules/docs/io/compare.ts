import { getSchema, type JSONContent } from '@tiptap/core'
import type { AttributeSpec, Schema } from '@tiptap/pm/model'
import { documentHeadersAndFooters, headerFooterParts } from '../header-footer'
import { headerFooterExtensions } from '../schema'
import { toPoints } from '../theme'
import { hexColor } from './docx-styles'

let headerFooterSchema: Schema | undefined

/** Attribute values that mean the same thing, written one way: left alignment is no alignment, colours in hex, sizes in points. */
function canonicalValue(key: string, value: unknown): unknown {
  if (key === 'textAlign' && value === 'left') return null
  if ((key === 'color' || key === 'backgroundColor') && typeof value === 'string') return hexColor(value) ?? value
  if (key === 'fontSize' && typeof value === 'string') {
    const points = toPoints(value)
    return points === undefined ? value : `${points}pt`
  }
  return value
}

/** Headers and footers through their own schema, as the body goes through its own. */
function canonicalHeaderFooter(value: unknown): unknown {
  const headerFooter = documentHeadersAndFooters(value)
  headerFooterSchema ??= getSchema(headerFooterExtensions)
  const schema = headerFooterSchema
  return {
    differentFirstPage: headerFooter.differentFirstPage,
    ...Object.fromEntries(headerFooterParts.map(part => {
      const content = headerFooter[part]
      return [part, content.length === 0 ? [] : canonical(schema.nodeFromJSON({ type: 'doc', content }).toJSON()).content]
    })),
  }
}

function canonicalAttrs(attrs: Readonly<Record<string, unknown>>): Record<string, unknown> {
  return Object.fromEntries(Object.entries(attrs).map(([key, value]) => [key, key === 'headerFooter' ? canonicalHeaderFooter(value) : canonicalValue(key, value)]))
}

function canonical(node: JSONContent): JSONContent {
  return {
    ...node,
    ...(node.attrs === undefined ? {} : { attrs: canonicalAttrs(node.attrs) }),
    ...(node.marks === undefined ? {} : { marks: node.marks.map(mark => mark.attrs === undefined ? mark : { ...mark, attrs: canonicalAttrs(mark.attrs) }) }),
    ...(node.content === undefined ? {} : { content: node.content.map(canonical) }),
  }
}

/**
 * Whether two documents hold the same content and formatting. Both go through the schema, which
 * fills every attribute default, joins neighbouring text with the same marks and orders marks,
 * so only real differences remain; the result is compared as JSON.
 */
export function sameDocument(schema: Schema, a: JSONContent, b: JSONContent): boolean {
  try {
    const text = (content: JSONContent): string => JSON.stringify(canonical(schema.nodeFromJSON(content).toJSON()))
    return text(a) === text(b)
  } catch {
    // A document the schema cannot build is not the same as anything.
    return false
  }
}

/** Each kind of thing a document uses: node and mark types, and attributes set away from their default (`paragraph.textAlign`). */
function kinds(schema: Schema, node: JSONContent, found: Set<string>): Set<string> {
  const add = (type: string, attrs: Readonly<Record<string, unknown>> | undefined, specs: Readonly<Record<string, AttributeSpec>> | undefined): void => {
    found.add(type)
    Object.entries(attrs ?? {}).forEach(([key, value]) => {
      if (JSON.stringify(canonicalValue(key, value)) !== JSON.stringify(canonicalValue(key, specs?.[key]?.default))) found.add(`${type}.${key}`)
    })
  }
  if (node.type !== undefined) add(node.type, node.attrs, schema.nodes[node.type]?.spec.attrs)
  node.marks?.forEach(mark => add(mark.type, mark.attrs, schema.marks[mark.type]?.spec.attrs))
  node.content?.forEach(child => kinds(schema, child, found))
  return found
}

/**
 * The kinds of thing `shown` uses that `reread` no longer does, sorted, e.g. `underline` or
 * `paragraph.textAlign`. Two saves that lose the same kinds give the same list, so a loss already
 * accepted is not asked about again, while a new kind of loss is.
 */
export function lostKinds(schema: Schema, shown: JSONContent, reread: JSONContent): readonly string[] {
  try {
    const used = (content: JSONContent): Set<string> => kinds(schema, schema.nodeFromJSON(content).toJSON(), new Set())
    const kept = used(reread)
    return [...used(shown)].filter(kind => !kept.has(kind)).sort()
  } catch {
    return ['unreadable']
  }
}
