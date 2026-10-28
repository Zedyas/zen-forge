import { getSchema, type JSONContent } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { documentExtensions } from '../schema'
import { lostKinds, sameDocument } from './compare'
import { readMarkdown, writeMarkdown } from './markdown'

const schema = getSchema(documentExtensions)
const doc = (...content: JSONContent[]): JSONContent => ({ type: 'doc', content })
const block = (type: string, text: string, attrs: Record<string, unknown> = {}): JSONContent => ({ type, attrs, content: [{ type: 'text', text }] })
const reopened = (content: JSONContent): JSONContent => readMarkdown(writeMarkdown(content)).content

describe('the check before saving over a file', () => {
  it('passes what Markdown keeps and catches what it would change', () => {
    const kept = doc(block('heading', 'Plan', { level: 1 }), block('paragraph', 'Text', { textAlign: 'left' }))
    expect(sameDocument(schema, kept, reopened(kept))).toBe(true)
    const withFooter = { ...doc(block('paragraph', 'Text')), attrs: { headerFooter: { footer: [block('paragraph', 'Page')] } } }
    for (const changed of [withFooter, doc(block('title', 'Plan')), doc(block('paragraph', 'Centred', { textAlign: 'center' })), doc(block('paragraph', 'x'), { type: 'pageBreak' }, block('paragraph', 'y'))]) {
      expect(sameDocument(schema, changed, reopened(changed))).toBe(false)
    }
  })

  it('names the kinds of thing a save loses, the same for more of the same loss', () => {
    const centred = doc(block('paragraph', 'Centred', { textAlign: 'center' }))
    const twice = doc(block('paragraph', 'One', { textAlign: 'center' }), block('paragraph', 'Two', { textAlign: 'center' }))
    const withTitle = doc(block('title', 'Plan'), block('paragraph', 'Centred', { textAlign: 'center' }))

    expect(lostKinds(schema, centred, reopened(centred))).toEqual(['paragraph.textAlign'])
    expect(lostKinds(schema, twice, reopened(twice))).toEqual(lostKinds(schema, centred, reopened(centred)))
    expect(lostKinds(schema, withTitle, reopened(withTitle))).toEqual(['paragraph.textAlign', 'title'])
  })
})
