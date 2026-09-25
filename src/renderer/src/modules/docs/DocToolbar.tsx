import { useState, type MouseEvent, type ReactNode } from 'react'
import type { Editor } from '@tiptap/core'
import { useEditorState } from '@tiptap/react'
import {
  Baseline,
  Bold,
  Highlighter,
  ImagePlus,
  Italic,
  List,
  ListIndentDecrease,
  ListIndentIncrease,
  ListOrdered,
  ListTodo,
  Minus,
  RemoveFormatting,
  Search,
  SeparatorHorizontal,
  Table,
  Underline,
} from 'lucide-react'
import { toast } from 'sonner'
import { TitleEssentials } from '../../ui/TitleSlot'
import { ColorTool, Toolbar, ToolButton, ToolSeparator, type ColorOption } from '../../ui/Toolbar'
import { clearFormatting, currentBlockStyle, currentStyle, indent, insertImage, insertPageBreak } from './doc-commands'
import { alignments, AlignMenu, FontMenu, LinkTool, MoreFormattingMenu, SizeMenu, SpacingMenu, StyleMenu, TableMenu, type AlignValue } from './DocMenus'
import type { StyleName } from './theme'

const textColors: readonly ColorOption[] = [
  { label: 'Automatic', value: undefined },
  { label: 'Black', value: '#000000' },
  { label: 'Dark red', value: '#c00000' },
  { label: 'Red', value: '#ff0000' },
  { label: 'Orange', value: '#e36c09' },
  { label: 'Green', value: '#00b050' },
  { label: 'Blue', value: '#0070c0' },
  { label: 'Purple', value: '#7030a0' },
]

/** Word's own highlight colours, so a highlight stays a highlight in Word. */
const highlightColors: readonly ColorOption[] = [
  { label: 'No highlight', value: undefined },
  { label: 'Yellow', value: '#ffff00' },
  { label: 'Green', value: '#00ff00' },
  { label: 'Cyan', value: '#00ffff' },
  { label: 'Pink', value: '#ff00ff' },
  { label: 'Grey', value: '#c0c0c0' },
]

interface FormatState {
  readonly style: StyleName
  readonly fontFamily: string
  readonly fontSize: number
  readonly styleFont: string
  readonly styleSize: number
  readonly styleLine: number
  readonly lineHeight: number
  readonly bold: boolean
  readonly italic: boolean
  readonly underline: boolean
  readonly strike: boolean
  readonly superscript: boolean
  readonly subscript: boolean
  readonly code: boolean
  readonly bulletList: boolean
  readonly orderedList: boolean
  readonly taskList: boolean
  readonly link: boolean
  readonly table: boolean
  readonly align: AlignValue
}

function points(value: unknown): number | undefined {
  const parsed = typeof value === 'string' ? Number.parseFloat(value) : Number.NaN
  return Number.isFinite(parsed) ? parsed : undefined
}

function formatState(editor: Editor): FormatState {
  const base = currentBlockStyle(editor)
  const textStyle = editor.getAttributes('textStyle')
  const family: unknown = textStyle['fontFamily']
  const line: unknown = editor.state.selection.$from.parent.attrs['lineHeight']
  return {
    style: currentStyle(editor),
    fontFamily: typeof family === 'string' && family !== '' ? family : base.fontFamily,
    fontSize: points(textStyle['fontSize']) ?? base.fontSize,
    styleFont: base.fontFamily,
    styleSize: base.fontSize,
    styleLine: base.lineHeight,
    lineHeight: typeof line === 'number' ? line : base.lineHeight,
    bold: editor.isActive('bold'),
    italic: editor.isActive('italic'),
    underline: editor.isActive('underline'),
    strike: editor.isActive('strike'),
    superscript: editor.isActive('superscript'),
    subscript: editor.isActive('subscript'),
    code: editor.isActive('code'),
    bulletList: editor.isActive('bulletList'),
    orderedList: editor.isActive('orderedList'),
    taskList: editor.isActive('taskList'),
    link: editor.isActive('link'),
    table: editor.isActive('table'),
    align: alignments.find(alignment => editor.isActive({ textAlign: alignment.value }))?.value ?? 'left',
  }
}

/** What the toolbar shows as on, read once per editor transaction. */
function useFormatState(editor: Editor): FormatState {
  return useEditorState({ editor, selector: ({ editor: current }) => formatState(current) })
}

/**
 * Toolbar buttons leave focus in the page, so the caret and selection stay put and typing carries
 * on straight after a click. Menu and popover triggers still take focus, as their popups need it.
 */
function KeepEditorFocus({ children }: { readonly children: ReactNode }) {
  const keepFocus = (event: MouseEvent): void => {
    const button = event.target instanceof Element ? event.target.closest('button') : null
    if (button !== null && !button.hasAttribute('aria-haspopup')) event.preventDefault()
  }
  return <div className="sumi-keep-focus" onMouseDown={keepFocus}>{children}</div>
}

function ColorTools({ editor }: { readonly editor: Editor }) {
  // The split colour buttons remember the last colour chosen in this window, like Word and Ledger.
  const [textColor, setTextColor] = useState<string | undefined>('#c00000')
  const [highlight, setHighlight] = useState<string | undefined>('#ffff00')
  const color = (value: string | undefined): void => {
    if (value === undefined) editor.chain().focus().unsetColor().run()
    else editor.chain().focus().setColor(value).run()
  }
  const background = (value: string | undefined): void => {
    if (value === undefined) editor.chain().focus().unsetBackgroundColor().run()
    else editor.chain().focus().setBackgroundColor(value).run()
  }
  return (
    <>
      <ColorTool icon={Baseline} label="Text colour" options={textColors} current={textColor} onApply={color} onChoose={value => {
        setTextColor(value)
        color(value)
      }} />
      <ColorTool icon={Highlighter} label="Highlight" options={highlightColors} current={highlight} onApply={background} onChoose={value => {
        setHighlight(value)
        background(value)
      }} />
    </>
  )
}

interface DocToolbarProps {
  readonly editor: Editor
  readonly findOpen: boolean
  onFind(): void
}

export function DocToolbar({ editor, findOpen, onFind }: DocToolbarProps) {
  const state = useFormatState(editor)
  const chain = () => editor.chain().focus()
  return (
    <KeepEditorFocus>
      <Toolbar label="Document tools">
        <StyleMenu editor={editor} current={state.style} />
        <FontMenu editor={editor} current={state.fontFamily} styleFont={state.styleFont} />
        <SizeMenu editor={editor} current={state.fontSize} styleSize={state.styleSize} />
        <ToolSeparator />
        <ToolButton icon={Bold} label="Bold" command="format.bold" pressed={state.bold} onClick={() => chain().toggleBold().run()} />
        <ToolButton icon={Italic} label="Italic" command="format.italic" pressed={state.italic} onClick={() => chain().toggleItalic().run()} />
        <ToolButton icon={Underline} label="Underline" command="format.underline" pressed={state.underline} onClick={() => chain().toggleUnderline().run()} />
        <ColorTools editor={editor} />
        <MoreFormattingMenu editor={editor} state={state} />
        <LinkTool editor={editor} active={state.link} />
        <ToolSeparator />
        <ToolButton icon={List} label="Bulleted list" shortcut="⇧⌘8" pressed={state.bulletList} onClick={() => chain().toggleBulletList().run()} />
        <ToolButton icon={ListOrdered} label="Numbered list" shortcut="⇧⌘7" pressed={state.orderedList} onClick={() => chain().toggleOrderedList().run()} />
        <ToolButton icon={ListTodo} label="Checklist" shortcut="⇧⌘9" pressed={state.taskList} onClick={() => chain().toggleTaskList().run()} />
        <ToolButton icon={ListIndentDecrease} label="Decrease indent" shortcut="⇧⇥" onClick={() => indent(editor, -1)} />
        <ToolButton icon={ListIndentIncrease} label="Increase indent" shortcut="⇥" onClick={() => indent(editor, 1)} />
        <ToolSeparator />
        <AlignMenu editor={editor} current={state.align} />
        <SpacingMenu editor={editor} current={state.lineHeight} styleLine={state.styleLine} />
        <ToolSeparator />
        <ToolButton icon={ImagePlus} label="Insert image…" onClick={() => void insertImage(editor).catch(() => toast.error('Could not insert the image'))} />
        <ToolButton icon={Table} label="Insert table" onClick={() => chain().insertTable({ rows: 3, cols: 3, withHeaderRow: false }).run()} />
        <TableMenu editor={editor} inTable={state.table} />
        <ToolButton icon={Minus} label="Horizontal line" onClick={() => chain().setHorizontalRule().run()} />
        <ToolButton icon={SeparatorHorizontal} label="Page break" shortcut="⌘↩" onClick={() => insertPageBreak(editor)} />
        <span className="toolbar-grow" />
        <ToolButton icon={RemoveFormatting} label="Clear formatting" command="format.clear" onClick={() => clearFormatting(editor)} />
        <ToolButton icon={Search} label="Find and replace" command="edit.find" pressed={findOpen} onClick={onFind} />
      </Toolbar>
    </KeepEditorFocus>
  )
}

/** The tools that stay in the title bar when the toolbar is hidden. */
export function DocEssentials({ editor }: { readonly editor: Editor }) {
  const state = useFormatState(editor)
  const chain = () => editor.chain().focus()
  return (
    <TitleEssentials>
      <KeepEditorFocus>
        <StyleMenu editor={editor} current={state.style} />
        <ToolButton icon={Bold} label="Bold" command="format.bold" pressed={state.bold} onClick={() => chain().toggleBold().run()} />
        <ToolButton icon={Italic} label="Italic" command="format.italic" pressed={state.italic} onClick={() => chain().toggleItalic().run()} />
        <ToolButton icon={Underline} label="Underline" command="format.underline" pressed={state.underline} onClick={() => chain().toggleUnderline().run()} />
        <ToolButton icon={List} label="Bulleted list" pressed={state.bulletList} onClick={() => chain().toggleBulletList().run()} />
        <ToolButton icon={ListOrdered} label="Numbered list" pressed={state.orderedList} onClick={() => chain().toggleOrderedList().run()} />
      </KeepEditorFocus>
    </TitleEssentials>
  )
}
