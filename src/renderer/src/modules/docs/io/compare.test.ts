import { getSchema, type JSONContent } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { documentExtensions } from '../schema'
import { sameDocument } from './compare'
import { readMarkdown, writeMarkdown } from './markdown'

const schema = getSchema(documentExtensions)
const doc = (...content: JSONContent[]): JSONContent => ({ type: 'doc', content })
const block = (type: string, text: string, attrs: Record<string, unknown> = {}): JSONContent => ({ type, attrs, content: [{ type: 'text', text }] })
const reopened = (content: JSONContent): JSONContent => readMarkdown(writeMarkdown(content)).content

describe('the check before saving over a file', () => {
  it('passes what Markdown keeps and catches what it would change', () => {
    const kept = doc(block('heading', 'Plan', { level: 1 }), block('paragraph', 'Text', { textAlign: 'left' }))
    expect(sameDocument(schema, kept, reopened(kept))).toBe(true)
    for (const changed of [doc(block('title', 'Plan')), doc(block('paragraph', 'Centred', { textAlign: 'center' })), doc(block('paragraph', 'x'), { type: 'pageBreak' }, block('paragraph', 'y'))]) {
      expect(sameDocument(schema, changed, reopened(changed))).toBe(false)
    }
  })
})
