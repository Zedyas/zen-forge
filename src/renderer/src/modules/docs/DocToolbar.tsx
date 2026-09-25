import { useState, type FormEvent, type MouseEvent, type ReactNode } from 'react'
import type { Editor } from '@tiptap/core'
import { useEditorState } from '@tiptap/react'
import { Menu } from '@base-ui/react/menu'
import { Popover } from '@base-ui/react/popover'
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  ChevronDown,
  ExternalLink,
  ImagePlus,
  Italic,
  Link2,
  List,
  ListOrdered,
  ListTodo,
  RemoveFormatting,
  Rows3,
  Search,
  SeparatorHorizontal,
  Strikethrough,
  Table,
  Underline,
} from 'lucide-react'
import { toast } from 'sonner'
import { Tip } from '../../ui/Tip'
import { TitleEssentials } from '../../ui/TitleSlot'
import { Toolbar, ToolButton, ToolSeparator } from '../../ui/Toolbar'
import { applyLink, clearFormatting, insertImage, insertPageBreak, openLink, removeLink } from './doc-commands'

type BlockStyle = 'paragraph' | 1 | 2 | 3

const blockStyles: ReadonlyArray<{ readonly value: BlockStyle; readonly label: string }> = [
  { value: 'paragraph', label: 'Normal text' },
  { value: 1, label: 'Heading 1' },
  { value: 2, label: 'Heading 2' },
  { value: 3, label: 'Heading 3' },
]

const alignments = [
  { value: 'left', label: 'Align left', icon: AlignLeft },
  { value: 'center', label: 'Align centre', icon: AlignCenter },
  { value: 'right', label: 'Align right', icon: AlignRight },
  { value: 'justify', label: 'Justify', icon: AlignJustify },
] as const

interface FormatState {
  readonly bold: boolean
  readonly italic: boolean
  readonly underline: boolean
  readonly strike: boolean
  readonly bulletList: boolean
  readonly orderedList: boolean
  readonly taskList: boolean
  readonly link: boolean
  readonly table: boolean
  readonly block: BlockStyle
  readonly align: (typeof alignments)[number]['value']
}

function formatState(editor: Editor): FormatState {
  return {
    bold: editor.isActive('bold'),
    italic: editor.isActive('italic'),
    underline: editor.isActive('underline'),
    strike: editor.isActive('strike'),
    bulletList: editor.isActive('bulletList'),
    orderedList: editor.isActive('orderedList'),
    taskList: editor.isActive('taskList'),
    link: editor.isActive('link'),
    table: editor.isActive('table'),
    block: ([1, 2, 3] as const).find(level => editor.isActive('heading', { level })) ?? 'paragraph',
    align: alignments.find(alignment => editor.isActive({ textAlign: alignment.value }))?.value ?? 'left',
  }
}

/** What the toolbar shows as on, read once per editor transaction. */
function useFormatState(editor: Editor): FormatState {
  return useEditorState({ editor, selector: ({ editor: current }) => formatState(current) })
}

function setBlockStyle(editor: Editor, style: BlockStyle): void {
  if (style === 'paragraph') editor.chain().focus().setParagraph().run()
  else editor.chain().focus().setHeading({ level: style }).run()
}

function BlockStyleMenu({ editor, current }: { readonly editor: Editor; readonly current: BlockStyle }) {
  const label = blockStyles.find(style => style.value === current)?.label ?? 'Normal text'
  return (
    <Menu.Root>
      <Tip label="Paragraph style">
        <Menu.Trigger className="tool is-label sumi-style-trigger" aria-label={`Paragraph style: ${label}`}>
          <span>{label}</span>
          <ChevronDown aria-hidden="true" size={11} strokeWidth={2.2} />
        </Menu.Trigger>
      </Tip>
      <Menu.Portal>
        <Menu.Positioner sideOffset={6} align="start">
          {/* The chosen style focuses the page again; returning focus to the trigger would take it away. */}
          <Menu.Popup className="menu-popup sumi-style-menu" finalFocus={false}>
            {blockStyles.map(style => (
              <Menu.Item
                key={style.value}
                className="menu-item"
                data-style={style.value}
                data-active={style.value === current ? '' : undefined}
                onClick={() => setBlockStyle(editor, style.value)}
              >
                <span>{style.label}</span>
              </Menu.Item>
            ))}
          </Menu.Popup>
        </Menu.Positioner>
      </Menu.Portal>
    </Menu.Root>
  )
}

function TableMenu({ editor, inTable }: { readonly editor: Editor; readonly inTable: boolean }) {
  const items: ReadonlyArray<readonly [string, () => boolean]> = [
    ['Add row above', () => editor.chain().focus().addRowBefore().run()],
    ['Add row below', () => editor.chain().focus().addRowAfter().run()],
    ['Delete row', () => editor.chain().focus().deleteRow().run()],
    ['Add column left', () => editor.chain().focus().addColumnBefore().run()],
    ['Add column right', () => editor.chain().focus().addColumnAfter().run()],
    ['Delete column', () => editor.chain().focus().deleteColumn().run()],
    ['Delete table', () => editor.chain().focus().deleteTable().run()],
  ]
  return (
    <Menu.Root>
      <Tip label={inTable ? 'Rows and columns' : 'Rows and columns: place the cursor in a table'}>
        <Menu.Trigger className="tool" aria-label="Rows and columns" disabled={!inTable}>
          <Rows3 aria-hidden="true" size={16} strokeWidth={1.7} />
          <ChevronDown aria-hidden="true" size={11} strokeWidth={2.2} />
        </Menu.Trigger>
      </Tip>
      <Menu.Portal>
        <Menu.Positioner sideOffset={6} align="start">
          <Menu.Popup className="menu-popup" finalFocus={false}>
            {items.map(([label, action], index) => (
              <div key={label}>
                {(index === 3 || index === 6) && <Menu.Separator className="menu-separator" />}
                <Menu.Item className="menu-item" onClick={action}><span>{label}</span></Menu.Item>
              </div>
            ))}
          </Menu.Popup>
        </Menu.Positioner>
      </Menu.Portal>
    </Menu.Root>
  )
}

function LinkTool({ editor, active }: { readonly editor: Editor; readonly active: boolean }) {
  const [open, setOpen] = useState(false)
  const [address, setAddress] = useState('')
  const href: unknown = editor.getAttributes('link')['href']
  const current = typeof href === 'string' ? href : undefined

  const submit = (event: FormEvent): void => {
    event.preventDefault()
    if (address.trim() === '') {
      if (active) removeLink(editor)
      setOpen(false)
      return
    }
    if (!applyLink(editor, address)) {
      toast.error('Use a web or email address', { description: 'For example example.com or name@example.com.' })
      return
    }
    setOpen(false)
  }

  return (
    <Popover.Root
      open={open}
      onOpenChange={next => {
        if (next) setAddress(current ?? '')
        setOpen(next)
      }}
    >
      <Tip label={active ? 'Edit link' : 'Link'}>
        <Popover.Trigger className="tool" aria-label="Link" aria-pressed={active}>
          <Link2 aria-hidden="true" size={16} strokeWidth={1.7} />
        </Popover.Trigger>
      </Tip>
      <Popover.Portal>
        <Popover.Positioner sideOffset={6} align="start">
          <Popover.Popup className="popover sumi-link-popover" finalFocus={false}>
            <form onSubmit={submit}>
              <input
                className="text-field"
                value={address}
                placeholder="Web or email address"
                aria-label="Link address"
                onChange={event => setAddress(event.currentTarget.value)}
              />
              <button type="submit" className="button is-primary">{active ? 'Update' : 'Link'}</button>
              {active && (
                <button type="button" className="button" onClick={() => {
                  removeLink(editor)
                  setOpen(false)
                }}>Remove</button>
              )}
              {current !== undefined && (
                <ToolButton icon={ExternalLink} label="Open in browser (⌘-click a link)" onClick={() => void openLink(current)} />
              )}
            </form>
          </Popover.Popup>
        </Popover.Positioner>
      </Popover.Portal>
    </Popover.Root>
  )
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
        <BlockStyleMenu editor={editor} current={state.block} />
        <ToolSeparator />
        <ToolButton icon={Bold} label="Bold" command="format.bold" pressed={state.bold} onClick={() => chain().toggleBold().run()} />
        <ToolButton icon={Italic} label="Italic" command="format.italic" pressed={state.italic} onClick={() => chain().toggleItalic().run()} />
        <ToolButton icon={Underline} label="Underline" command="format.underline" pressed={state.underline} onClick={() => chain().toggleUnderline().run()} />
        <ToolButton icon={Strikethrough} label="Strikethrough" shortcut="⇧⌘X" pressed={state.strike} onClick={() => chain().toggleStrike().run()} />
        <LinkTool editor={editor} active={state.link} />
        <ToolSeparator />
        <ToolButton icon={List} label="Bulleted list" shortcut="⇧⌘8" pressed={state.bulletList} onClick={() => chain().toggleBulletList().run()} />
        <ToolButton icon={ListOrdered} label="Numbered list" shortcut="⇧⌘7" pressed={state.orderedList} onClick={() => chain().toggleOrderedList().run()} />
        <ToolButton icon={ListTodo} label="Checklist" shortcut="⇧⌘9" pressed={state.taskList} onClick={() => chain().toggleTaskList().run()} />
        <ToolSeparator />
        {alignments.map(alignment => (
          <ToolButton
            key={alignment.value}
            icon={alignment.icon}
            label={alignment.label}
            pressed={state.align === alignment.value}
            onClick={() => chain().setTextAlign(alignment.value).run()}
          />
        ))}
        <ToolSeparator />
        <ToolButton icon={ImagePlus} label="Insert image…" onClick={() => void insertImage(editor).catch(() => toast.error('Could not insert the image'))} />
        <ToolButton icon={Table} label="Insert table" onClick={() => chain().insertTable({ rows: 3, cols: 3, withHeaderRow: false }).run()} />
        <TableMenu editor={editor} inTable={state.table} />
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
        <ToolButton icon={Bold} label="Bold" command="format.bold" pressed={state.bold} onClick={() => chain().toggleBold().run()} />
        <ToolButton icon={Italic} label="Italic" command="format.italic" pressed={state.italic} onClick={() => chain().toggleItalic().run()} />
        <ToolButton icon={Underline} label="Underline" command="format.underline" pressed={state.underline} onClick={() => chain().toggleUnderline().run()} />
        <ToolButton icon={List} label="Bulleted list" pressed={state.bulletList} onClick={() => chain().toggleBulletList().run()} />
        <ToolButton icon={ListOrdered} label="Numbered list" pressed={state.orderedList} onClick={() => chain().toggleOrderedList().run()} />
      </KeepEditorFocus>
    </TitleEssentials>
  )
}
