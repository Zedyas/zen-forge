// @vitest-environment happy-dom
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { findState, FindHighlight, replaceAll, replaceCurrent, setFindQuery } from './find'
import { documentExtensions } from './schema'

function editorWith(value: string): Editor {
  return new Editor({
    extensions: [...documentExtensions, FindHighlight],
    content: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: value }] }] },
  })
}

describe('find and replace', () => {
  it('moves on to the next match when the replacement holds the search text, and does nothing once none is left', () => {
    const editor = editorWith('test one, test two')
    setFindQuery(editor, 'test', false)
    for (let count = 0; count < 4; count += 1) replaceCurrent(editor, 'tests')
    const text = editor.state.doc.textContent
    const matches = findState(editor.state)?.matches.length
    editor.destroy()

    expect(text).toBe('tests one, tests two')
    expect(matches).toBe(0)
  })

  it('does not replace the text a replacement inserted when only one match was there', () => {
    const editor = editorWith('test')
    setFindQuery(editor, 'test', false)
    replaceCurrent(editor, 'tests')
    replaceCurrent(editor, 'tests')
    const text = editor.state.doc.textContent
    editor.destroy()

    expect(text).toBe('tests')
  })

  it('finds the search text in the replaced text again once the search changes', () => {
    const editor = editorWith('test')
    setFindQuery(editor, 'test', false)
    replaceCurrent(editor, 'tests')
    setFindQuery(editor, 'tes', false)
    const matches = findState(editor.state)?.matches.length
    editor.destroy()

    expect(matches).toBe(1)
  })

  it('leaves the text Replace All inserted out of the matches', () => {
    const editor = editorWith('test one, test two')
    setFindQuery(editor, 'test', false)
    replaceAll(editor, 'tests')
    const matches = findState(editor.state)?.matches.length
    editor.destroy()

    expect(matches).toBe(0)
  })
})
