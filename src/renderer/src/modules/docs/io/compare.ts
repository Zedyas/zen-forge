import type { JSONContent } from '@tiptap/core'
import type { Schema } from '@tiptap/pm/model'
import { toPoints } from '../theme'
import { hexColor } from './docx-styles'

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

function canonicalAttrs(attrs: Readonly<Record<string, unknown>>): Record<string, unknown> {
  return Object.fromEntries(Object.entries(attrs).map(([key, value]) => [key, canonicalValue(key, value)]))
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
