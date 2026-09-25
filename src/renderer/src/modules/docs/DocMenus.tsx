import { useState, type CSSProperties, type FormEvent, type ReactNode } from 'react'
import type { Editor } from '@tiptap/core'
import { Menu } from '@base-ui/react/menu'
import { Popover } from '@base-ui/react/popover'
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Check,
  ChevronDown,
  Code,
  Ellipsis,
  ExternalLink,
  Link2,
  Rows3,
  Strikethrough,
  Subscript,
  Superscript,
  UnfoldVertical,
  type LucideIcon,
} from 'lucide-react'
import { toast } from 'sonner'
import { Tip } from '../../ui/Tip'
import { ToolButton } from '../../ui/Toolbar'
import { applyLink, openLink, removeLink, setBlockStyle, setParagraphFormat } from './doc-commands'
import { documentTheme, fontFamilies, fontSizes, lineSpacings, styleLabels, styleNames, type StyleName } from './theme'

/** A toolbar menu: a trigger with a tooltip and a chevron, and a popup that leaves focus to the page. */
function ToolMenu({ label, trigger, wide, disabled, children }: {
  readonly label: string
  readonly trigger: ReactNode
  readonly wide?: boolean
  readonly disabled?: boolean
  readonly children: ReactNode
}) {
  return (
    <Menu.Root>
      <Tip label={label}>
        <Menu.Trigger className={`tool${wide === true ? ' is-label sumi-menu-trigger' : ''}`} aria-label={label} disabled={disabled}>
          {trigger}
          <ChevronDown aria-hidden="true" size={11} strokeWidth={2.2} />
        </Menu.Trigger>
      </Tip>
      <Menu.Portal>
        <Menu.Positioner sideOffset={6} align="start">
          {/* Each item focuses the page again; returning focus to the trigger would take it away. */}
          <Menu.Popup className="menu-popup sumi-menu" finalFocus={false}>{children}</Menu.Popup>
        </Menu.Positioner>
      </Menu.Portal>
    </Menu.Root>
  )
}

function Item({ active, onClick, children, style }: {
  readonly active?: boolean
  readonly onClick: () => void
  readonly children: ReactNode
  readonly style?: CSSProperties
}) {
  return (
    <Menu.Item className="menu-item" data-active={active === true ? '' : undefined} onClick={onClick}>
      <span style={style}>{children}</span>
      {active === true && <Check aria-hidden="true" size={14} />}
    </Menu.Item>
  )
}

/** Each style previewed in the document's own fonts, at a size that fits the menu. */
export function StyleMenu({ editor, current }: { readonly editor: Editor; readonly current: StyleName }) {
  const theme = documentTheme(editor.state.doc.attrs['theme'])
  return (
    <ToolMenu label="Paragraph style" wide trigger={<span>{styleLabels[current]}</span>}>
      {styleNames.map(name => {
        const style = theme[name]
        const preview: CSSProperties = {
          fontFamily: `"${style.fontFamily}"`,
          fontSize: `${Math.min(20, Math.max(13, style.fontSize * 0.9))}px`,
          fontWeight: style.bold ? 700 : 400,
          fontStyle: style.italic ? 'italic' : 'normal',
        }
        return <Item key={name} active={name === current} style={preview} onClick={() => setBlockStyle(editor, name)}>{styleLabels[name]}</Item>
      })}
    </ToolMenu>
  )
}

/** Fonts used in the document that are not in the usual list, so they can be picked again. */
function documentFonts(editor: Editor): string[] {
  const fonts = new Set<string>()
  editor.state.doc.descendants(node => {
    for (const mark of node.marks) {
      const family: unknown = mark.attrs['fontFamily']
      if (typeof family === 'string' && family !== '') fonts.add(family)
    }
  })
  const listed: readonly string[] = fontFamilies
  return [...fonts].filter(font => !listed.includes(font)).sort()
}

export function FontMenu({ editor, current, styleFont }: { readonly editor: Editor; readonly current: string; readonly styleFont: string }) {
  const choose = (font: string): void => {
    // The style's own font is no formatting at all, so the text follows the style if it changes.
    if (font === styleFont) editor.chain().focus().unsetFontFamily().run()
    else editor.chain().focus().setFontFamily(font).run()
  }
  const fonts = [...new Set([styleFont, ...fontFamilies])]
  const extra = documentFonts(editor).filter(font => font !== styleFont)
  return (
    <ToolMenu label="Font" wide trigger={<span className="sumi-font-name">{current}</span>}>
      {fonts.map(font => <Item key={font} active={font === current} style={{ fontFamily: `"${font}"` }} onClick={() => choose(font)}>{font}</Item>)}
      {extra.length > 0 && <Menu.Separator className="menu-separator" />}
      {extra.map(font => <Item key={font} active={font === current} style={{ fontFamily: `"${font}"` }} onClick={() => choose(font)}>{font}</Item>)}
    </ToolMenu>
  )
}

export function SizeMenu({ editor, current, styleSize }: { readonly editor: Editor; readonly current: number; readonly styleSize: number }) {
  const choose = (size: number): void => {
    if (size === styleSize) editor.chain().focus().unsetFontSize().run()
    else editor.chain().focus().setFontSize(`${size}pt`).run()
  }
  const sizes = [...new Set([...fontSizes, styleSize])].sort((a, b) => a - b)
  return (
    <ToolMenu label="Font size" wide trigger={<span className="sumi-size-value">{current}</span>}>
      {sizes.map(size => <Item key={size} active={size === current} onClick={() => choose(size)}>{size}</Item>)}
    </ToolMenu>
  )
}

interface MoreState {
  readonly strike: boolean
  readonly superscript: boolean
  readonly subscript: boolean
  readonly code: boolean
}

/** Rarer character formats, one menu away. */
export function MoreFormattingMenu({ editor, state }: { readonly editor: Editor; readonly state: MoreState }) {
  const items: ReadonlyArray<readonly [LucideIcon, string, string, boolean, () => boolean]> = [
    [Strikethrough, 'Strikethrough', '⇧⌘X', state.strike, () => editor.chain().focus().toggleStrike().run()],
    [Superscript, 'Superscript', '⌘.', state.superscript, () => editor.chain().focus().toggleSuperscript().run()],
    [Subscript, 'Subscript', '⌘,', state.subscript, () => editor.chain().focus().toggleSubscript().run()],
    [Code, 'Code', '⌘E', state.code, () => editor.chain().focus().toggleCode().run()],
  ]
  return (
    <ToolMenu label="More text formatting" trigger={<Ellipsis aria-hidden="true" size={16} strokeWidth={1.7} />}>
      {items.map(([Icon, label, shortcut, active, action]) => (
        <Menu.Item key={label} className="menu-item" data-active={active ? '' : undefined} onClick={action}>
          <Icon aria-hidden="true" size={16} strokeWidth={1.7} />
          <span>{label}</span>
          <kbd>{shortcut}</kbd>
        </Menu.Item>
      ))}
    </ToolMenu>
  )
}

export const alignments = [
  { value: 'left', label: 'Align left', icon: AlignLeft, shortcut: '⇧⌘L' },
  { value: 'center', label: 'Align centre', icon: AlignCenter, shortcut: '⇧⌘E' },
  { value: 'right', label: 'Align right', icon: AlignRight, shortcut: '⇧⌘R' },
  { value: 'justify', label: 'Justify', icon: AlignJustify, shortcut: '⇧⌘J' },
] as const

export type AlignValue = (typeof alignments)[number]['value']

export function AlignMenu({ editor, current }: { readonly editor: Editor; readonly current: AlignValue }) {
  const Icon = alignments.find(alignment => alignment.value === current)?.icon ?? AlignLeft
  return (
    <ToolMenu label="Alignment" trigger={<Icon aria-hidden="true" size={16} strokeWidth={1.7} />}>
      {alignments.map(alignment => (
        <Menu.Item
          key={alignment.value}
          className="menu-item"
          data-active={alignment.value === current ? '' : undefined}
          onClick={() => editor.chain().focus().setTextAlign(alignment.value).run()}
        >
          <alignment.icon aria-hidden="true" size={16} strokeWidth={1.7} />
          <span>{alignment.label}</span>
          <kbd>{alignment.shortcut}</kbd>
        </Menu.Item>
      ))}
    </ToolMenu>
  )
}

export function SpacingMenu({ editor, current, styleLine }: { readonly editor: Editor; readonly current: number; readonly styleLine: number }) {
  return (
    <ToolMenu label="Line spacing" trigger={<UnfoldVertical aria-hidden="true" size={16} strokeWidth={1.7} />}>
      {[...new Set([...lineSpacings, styleLine])].sort((a, b) => a - b).map(spacing => (
        <Item key={spacing} active={spacing === current} onClick={() => setParagraphFormat(editor, { lineHeight: spacing === styleLine ? null : spacing })}>
          {spacing === 1 ? 'Single' : spacing === 2 ? 'Double' : String(spacing)}
        </Item>
      ))}
      <Menu.Separator className="menu-separator" />
      <Item onClick={() => setParagraphFormat(editor, { spaceBefore: null, spaceAfter: null })}>Reset space before and after</Item>
    </ToolMenu>
  )
}

export function TableMenu({ editor, inTable }: { readonly editor: Editor; readonly inTable: boolean }) {
  const chain = () => editor.chain().focus()
  const groups: ReadonlyArray<ReadonlyArray<readonly [string, () => boolean, boolean]>> = [
    [
      ['Add row above', () => chain().addRowBefore().run(), true],
      ['Add row below', () => chain().addRowAfter().run(), true],
      ['Delete row', () => chain().deleteRow().run(), true],
    ],
    [
      ['Add column left', () => chain().addColumnBefore().run(), true],
      ['Add column right', () => chain().addColumnAfter().run(), true],
      ['Delete column', () => chain().deleteColumn().run(), true],
    ],
    [
      ['Header row', () => chain().toggleHeaderRow().run(), true],
      ['Merge cells', () => chain().mergeCells().run(), editor.can().mergeCells()],
      ['Split cell', () => chain().splitCell().run(), editor.can().splitCell()],
    ],
    [['Delete table', () => chain().deleteTable().run(), true]],
  ]
  return (
    <ToolMenu label={inTable ? 'Rows and columns' : 'Rows and columns: place the cursor in a table'} disabled={!inTable} trigger={<Rows3 aria-hidden="true" size={16} strokeWidth={1.7} />}>
      {groups.map((group, index) => (
        <div key={group[0]?.[0]}>
          {index > 0 && <Menu.Separator className="menu-separator" />}
          {group.map(([label, action, enabled]) => (
            <Menu.Item key={label} className="menu-item" disabled={!enabled} onClick={action}><span>{label}</span></Menu.Item>
          ))}
        </div>
      ))}
    </ToolMenu>
  )
}

export function LinkTool({ editor, active }: { readonly editor: Editor; readonly active: boolean }) {
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
