// @vitest-environment happy-dom
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { FindHighlight, replaceCurrent, setFindQuery } from './find'
import { documentExtensions } from './schema'

describe('find and replace', () => {
  it('moves on to the next match when the replacement holds the search text', () => {
    const editor = new Editor({
      extensions: [...documentExtensions, FindHighlight],
      content: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'test one, test two' }] }] },
    })
    setFindQuery(editor, 'test', false)
    replaceCurrent(editor, 'tests')
    replaceCurrent(editor, 'tests')
    const text = editor.state.doc.textContent
    editor.destroy()

    expect(text).toBe('tests one, tests two')
  })
})
