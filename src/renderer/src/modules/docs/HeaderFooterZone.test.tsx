// @vitest-environment happy-dom
import { Editor, type JSONContent } from '@tiptap/core'
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it } from 'vitest'
import { editPart, setHeadersAndFooters } from './doc-commands'
import { documentHeadersAndFooters, hasPageNumber, withPageNumber, type HeadersAndFooters } from './header-footer'
import { HeaderFooterZone } from './HeaderFooterZone'
import { documentExtensions } from './schema'
import { defaultTheme } from './theme'

Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })

const cleanups: Array<() => void> = []
afterEach(() => cleanups.splice(0).forEach(cleanup => cleanup()))

const line = (text: string): JSONContent => ({ type: 'paragraph', content: [{ type: 'text', text }] })

/** The bottom margin of a page whose headers and footers are `change`. */
function footerZone(change: Partial<HeadersAndFooters>): { readonly editor: Editor; readonly host: HTMLElement } {
  const editor = new Editor({ extensions: documentExtensions, content: { type: 'doc', content: [line('Body')] } })
  setHeadersAndFooters(editor, change)
  const host = document.createElement('div')
  document.body.append(host)
  const root = createRoot(host)
  act(() => root.render(<HeaderFooterZone editor={editor} documentId="doc" edge="bottom" base={defaultTheme.normal} style={{}} />))
  cleanups.push(() => {
    act(() => root.unmount())
    act(() => editPart('doc', undefined))
    editor.destroy()
    host.remove()
  })
  return { editor, host }
}

function footers(editor: Editor): HeadersAndFooters {
  return documentHeadersAndFooters(editor.state.doc.attrs['headerFooter'])
}

describe('header and footer zone', () => {
  it('takes in a change made while it is open, so typing does not write the old footer back', () => {
    const { editor, host } = footerZone({ footer: [line('Draft')] })
    act(() => editPart('doc', 'footer'))
    act(() => setHeadersAndFooters(editor, { footer: withPageNumber(footers(editor).footer, true) }))

    const dom = host.querySelector('.sumi-part-editor')
    const partEditor = dom !== null && 'editor' in dom && dom.editor instanceof Editor ? dom.editor : undefined
    act(() => partEditor?.chain().focus('start').insertContent('Final ').run())

    expect(footers(editor).footer[0]).toMatchObject(line('Final Draft'))
    expect(hasPageNumber(footers(editor).footer)).toBe(true)
  })

  it('shows and edits the first page footer while the first page has its own', () => {
    const { host } = footerZone({ footer: [line('Draft')], firstFooter: [line('Cover')], differentFirstPage: true })
    expect(host.textContent).toBe('Cover')

    act(() => host.querySelector('.sumi-part')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })))
    expect(host.querySelector('.sumi-part-editor')?.getAttribute('aria-label')).toBe('First page footer')
    expect(host.querySelector('.sumi-part-editor')?.textContent).toBe('Cover')
  })
})
