import { useEffect, useLayoutEffect, useRef, type CSSProperties, type MouseEvent } from 'react'
import { getSchema, type Editor } from '@tiptap/core'
import { DOMSerializer, Node as ProseMirrorNode } from '@tiptap/pm/model'
import { EditorContent, useEditor, useEditorState } from '@tiptap/react'
import { AlignCenter, AlignLeft, AlignRight, Bold, Check, Files, Hash, Italic, Underline } from 'lucide-react'
import { ToolButton, ToolSeparator } from '../../ui/Toolbar'
import { editPart, setActivePart, setHeadersAndFooters, usePartEditing } from './doc-commands'
import { FontMenu, SizeMenu } from './DocMenus'
import { documentHeadersAndFooters, partLabels, trimmedPart, type HeaderFooterContent, type HeaderFooterPart } from './header-footer'
import { headerFooterExtensions } from './schema'
import { toPoints, type BlockStyle } from './theme'

const schema = getSchema(headerFooterExtensions)

/** A header or footer as it shows when not being edited, drawn by its schema like the page itself. */
function PartView({ content }: { readonly content: HeaderFooterContent }) {
  const view = useRef<HTMLDivElement>(null)
  useLayoutEffect(() => {
    if (view.current === null) return
    const doc = content.length === 0 ? undefined : ProseMirrorNode.fromJSON(schema, { type: 'doc', content })
    view.current.replaceChildren(doc === undefined ? '' : DOMSerializer.fromSchema(schema).serializeFragment(doc.content))
  }, [content])
  return <div ref={view} className="sumi-part-view" />
}

function PartToolbar({ part, editor, base, onDone }: { readonly part: HeaderFooterPart; readonly editor: Editor; readonly base: BlockStyle; onDone(): void }) {
  const state = useEditorState({
    editor,
    selector: ({ editor: current }) => {
      const family: unknown = current.getAttributes('textStyle')['fontFamily']
      return {
        bold: current.isActive('bold'),
        italic: current.isActive('italic'),
        underline: current.isActive('underline'),
        align: (['center', 'right'] as const).find(align => current.isActive({ textAlign: align })) ?? 'left',
        fontFamily: typeof family === 'string' && family !== '' ? family : base.fontFamily,
        fontSize: toPoints(current.getAttributes('textStyle')['fontSize']) ?? base.fontSize,
      }
    },
  })
  const chain = () => editor.chain().focus()
  const field = (kind: 'page' | 'total') => chain().insertContent({ type: 'pageField', attrs: { kind } }).run()
  // Buttons leave focus in the header, so typing carries on after a click.
  const keepFocus = (event: MouseEvent): void => {
    const button = event.target instanceof Element ? event.target.closest('button') : null
    if (button !== null && !button.hasAttribute('aria-haspopup')) event.preventDefault()
  }
  return (
    <div className="sumi-part-toolbar" role="toolbar" aria-label={`${partLabels[part]} tools`} onMouseDown={keepFocus}>
      <span className="sumi-part-label">{partLabels[part]}</span>
      <FontMenu editor={editor} current={state.fontFamily} styleFont={base.fontFamily} />
      <SizeMenu editor={editor} current={state.fontSize} styleSize={base.fontSize} />
      <ToolSeparator />
      <ToolButton icon={Bold} label="Bold" command="format.bold" pressed={state.bold} onClick={() => chain().toggleBold().run()} />
      <ToolButton icon={Italic} label="Italic" command="format.italic" pressed={state.italic} onClick={() => chain().toggleItalic().run()} />
      <ToolButton icon={Underline} label="Underline" command="format.underline" pressed={state.underline} onClick={() => chain().toggleUnderline().run()} />
      <ToolSeparator />
      <ToolButton icon={AlignLeft} label="Align left" pressed={state.align === 'left'} onClick={() => chain().setTextAlign('left').run()} />
      <ToolButton icon={AlignCenter} label="Align centre" pressed={state.align === 'center'} onClick={() => chain().setTextAlign('center').run()} />
      <ToolButton icon={AlignRight} label="Align right" pressed={state.align === 'right'} onClick={() => chain().setTextAlign('right').run()} />
      <ToolSeparator />
      <ToolButton icon={Hash} label="Insert page number" onClick={() => field('page')} />
      <ToolButton icon={Files} label="Insert page count" onClick={() => field('total')} />
      <ToolSeparator />
      <ToolButton icon={Check} label="Done (Escape)" onClick={onDone} />
    </div>
  )
}

/** The small editor for one header or footer. Each change goes into the document at once, so saving never misses it. */
function PartEditor({ editor, part, base }: { readonly editor: Editor; readonly part: HeaderFooterPart; readonly base: BlockStyle }) {
  const close = (): void => {
    usePartEditing.setState({ editing: undefined })
    editor.commands.focus()
  }
  const partEditor = useEditor({
    extensions: headerFooterExtensions,
    content: (() => {
      const content = documentHeadersAndFooters(editor.state.doc.attrs['headerFooter'])[part]
      return { type: 'doc', content: content.length > 0 ? [...content] : [{ type: 'paragraph' }] }
    })(),
    autofocus: 'end',
    editorProps: {
      attributes: { class: 'sumi-part-editor', spellcheck: 'true', 'aria-label': partLabels[part] },
      handleKeyDown: (_view, event) => {
        if (event.key !== 'Escape') return false
        close()
        return true
      },
    },
    onUpdate: ({ editor: current }) => setHeadersAndFooters(editor, { [part]: trimmedPart(current.getJSON().content ?? []) }),
  }, [part])

  useEffect(() => setActivePart(partEditor), [partEditor])
  // Clicking into the page ends editing, as in Word.
  useEffect(() => {
    const end = (): void => usePartEditing.setState({ editing: undefined })
    editor.on('focus', end)
    return () => {
      editor.off('focus', end)
    }
  }, [editor])

  return (
    <>
      <PartToolbar part={part} editor={partEditor} base={base} onDone={close} />
      <EditorContent editor={partEditor} className="sumi-part-content" />
    </>
  )
}

const hints: Readonly<Record<'top' | 'bottom', string>> = { top: 'Double-click to add a header', bottom: 'Double-click to add a footer' }

/**
 * The header or footer in the page's top or bottom margin. The page on screen is one long sheet,
 * so its top shows the first page's header and its bottom the footer most pages have.
 */
export function HeaderFooterZone({ editor, documentId, edge, base, style }: {
  readonly editor: Editor
  readonly documentId: string
  readonly edge: 'top' | 'bottom'
  readonly base: BlockStyle
  /** Where the margin is on the page. */
  readonly style: CSSProperties
}) {
  const headerFooter = useEditorState({ editor, selector: ({ editor: current }) => documentHeadersAndFooters(current.state.doc.attrs['headerFooter']) })
  const editing = usePartEditing(state => state.editing)
  const parts: readonly HeaderFooterPart[] = edge === 'top' ? ['header', 'firstHeader'] : ['footer', 'firstFooter']
  const editingHere = editing !== undefined && editing.documentId === documentId && parts.includes(editing.part) ? editing.part : undefined
  const shown: HeaderFooterPart = edge === 'top' && headerFooter.differentFirstPage ? 'firstHeader' : parts[0] ?? 'header'
  const content = headerFooter[shown]
  return (
    <div
      className={`sumi-part is-${edge}`}
      style={style}
      data-editing={editingHere === undefined ? undefined : ''}
      data-hint={content.length === 0 && editingHere === undefined ? hints[edge] : undefined}
      onDoubleClick={editingHere === undefined ? () => editPart(documentId, shown) : undefined}
    >
      {editingHere === undefined ? <PartView content={content} /> : <PartEditor key={editingHere} editor={editor} part={editingHere} base={base} />}
    </div>
  )
}
