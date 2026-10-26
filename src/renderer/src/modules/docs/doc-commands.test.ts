// @vitest-environment happy-dom
import { Editor, getSchema } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { insertTableOfContents } from './doc-commands'
import { sameDocument } from './io/compare'
import { writeDocx } from './io/docx-export'
import { readDocx } from './io/docx-read'
import { documentExtensions } from './schema'

describe('table of contents', () => {
  it('goes after the list the caret is in, so a Word file keeps the list whole', async () => {
    const item = (text: string) => ({ type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text }] }] })
    const editor = new Editor({
      extensions: documentExtensions,
      content: { type: 'doc', content: [{ type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Plan' }] }, { type: 'bulletList', content: [item('One'), item('Two')] }] },
    })
    // In the first item's text.
    editor.commands.setTextSelection(editor.state.doc.resolve(0).posAtIndex(1) + 4)
    insertTableOfContents(editor)
    const content = editor.getJSON()
    editor.destroy()

    expect(content.content?.map(node => node.type)).toEqual(['heading', 'bulletList', 'tableOfContents'])
    const read = readDocx((await writeDocx(content)).bytes)
    expect(read.findings).toEqual([])
    expect(sameDocument(getSchema(documentExtensions), read.content, content)).toBe(true)
  })
})
