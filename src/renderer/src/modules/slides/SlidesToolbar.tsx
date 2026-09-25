import { Menu } from '@base-ui/react/menu'
import { Popover } from '@base-ui/react/popover'
import {
  AlignCenter,
  AlignCenterHorizontal,
  AlignCenterVertical,
  AlignEndHorizontal,
  AlignEndVertical,
  AlignLeft,
  AlignRight,
  AlignStartHorizontal,
  AlignStartVertical,
  ArrowBigRight,
  Baseline,
  Bold,
  BringToFront,
  ChevronDown,
  Circle,
  Highlighter,
  ImagePlus,
  Italic,
  Layers,
  List,
  ListOrdered,
  Minus,
  MoveRight,
  PaintBucket,
  Play,
  Plus,
  Redo2,
  SendToBack,
  Shapes,
  Square,
  SquareRoundCorner,
  Table,
  Triangle,
  Type,
  Underline,
  Undo2,
  type LucideIcon,
} from 'lucide-react'
import { useState, type ReactNode } from 'react'
import { toast } from 'sonner'
import type { CommandId } from '@shared/commands'
import { icon, tinyIcon } from '../../ui/icons'
import { commandShortcut, Tip } from '../../ui/Tip'
import { ColorTool, Toolbar, ToolButton, ToolSeparator } from '../../ui/Toolbar'
import { TitleEssentials } from '../../ui/TitleSlot'
import { useViewStore } from '../../app/view-store'
import { applyRunStyle, insertTable, setAlign, setFill, toggleList, toggleRunStyle } from './format-actions'
import { allRuns, holdsText, newShape, newTextBox, slideLayouts, type AlignEdge, type BasicShape, type HorizontalAlign, type TextRun } from './model'
import { fillColors, fontChoices, fontSizes, highlightColors, textColors } from './palette'
import { runSlidesCommand } from './slide-commands'
import { addElement, addSlide, alignSelection, arrangeSelection, insertPicture } from './slides-actions'
import { selectedElements, type ReadySlides } from './slides-store'

const shapes: ReadonlyArray<{ readonly shape: BasicShape; readonly label: string; readonly icon: LucideIcon }> = [
  { shape: 'rect', label: 'Rectangle', icon: Square },
  { shape: 'roundRect', label: 'Rounded rectangle', icon: SquareRoundCorner },
  { shape: 'ellipse', label: 'Ellipse', icon: Circle },
  { shape: 'triangle', label: 'Triangle', icon: Triangle },
  { shape: 'rightArrow', label: 'Arrow', icon: ArrowBigRight },
  { shape: 'line', label: 'Line', icon: Minus },
  { shape: 'arrow', label: 'Line with arrowhead', icon: MoveRight },
]

const alignments: ReadonlyArray<{ readonly value: HorizontalAlign; readonly label: string; readonly icon: LucideIcon }> = [
  { value: 'left', label: 'Align text left', icon: AlignLeft },
  { value: 'center', label: 'Align text centre', icon: AlignCenter },
  { value: 'right', label: 'Align text right', icon: AlignRight },
]

export const objectAlignments: ReadonlyArray<{ readonly edge: AlignEdge; readonly label: string; readonly icon: LucideIcon }> = [
  { edge: 'left', label: 'Align left edges', icon: AlignStartVertical },
  { edge: 'center', label: 'Align centres', icon: AlignCenterVertical },
  { edge: 'right', label: 'Align right edges', icon: AlignEndVertical },
  { edge: 'top', label: 'Align top edges', icon: AlignStartHorizontal },
  { edge: 'middle', label: 'Align middles', icon: AlignCenterHorizontal },
  { edge: 'bottom', label: 'Align bottom edges', icon: AlignEndHorizontal },
]

function run(command: CommandId): void {
  void runSlidesCommand(command)
}

function reportFailure(title: string) {
  return (error: unknown): void => {
    toast.error(title, { description: error instanceof Error ? error.message : undefined })
  }
}

interface ToolProps {
  readonly documentId: string
  readonly document: ReadySlides
}

/**
 * A toolbar button that opens a menu; the chevron says so. Choosing an item leaves focus where the
 * action puts it (the text being typed, or the slide) rather than on the button, so arrow keys then
 * nudge the new object instead of reopening the menu.
 */
function MenuTool({ icon: Icon, label, text, shortcut, disabled, children }: {
  readonly icon?: LucideIcon
  readonly label: string
  readonly text?: string
  readonly shortcut?: string
  readonly disabled?: boolean
  readonly children: ReactNode
}) {
  return (
    <Menu.Root>
      <Tip label={label} shortcut={shortcut}>
        <Menu.Trigger className={`tool${text === undefined ? '' : ' is-label menu-label'}`} aria-label={label} disabled={disabled}>
          {Icon !== undefined && <Icon {...icon} />}
          {text !== undefined && <span>{text}</span>}
          <ChevronDown {...tinyIcon} />
        </Menu.Trigger>
      </Tip>
      <Menu.Portal>
        <Menu.Positioner sideOffset={6} align="start">
          <Menu.Popup className="menu-popup slides-menu" finalFocus={false}>{children}</Menu.Popup>
        </Menu.Positioner>
      </Menu.Portal>
    </Menu.Root>
  )
}

function NewSlideMenu({ documentId }: { readonly documentId: string }) {
  return (
    <MenuTool icon={Plus} label="New slide" shortcut={commandShortcut('slide.new')}>
      {slideLayouts.map(layout => (
        <Menu.Item key={layout.id} className="menu-item" onClick={() => addSlide(documentId, layout.id)}><span>{layout.label}</span></Menu.Item>
      ))}
    </MenuTool>
  )
}

function ShapeMenu({ documentId, document }: ToolProps) {
  return (
    <MenuTool icon={Shapes} label="Shape">
      {shapes.map(({ shape, label, icon: Icon }) => (
        <Menu.Item key={shape} className="menu-item" onClick={() => addElement(documentId, newShape(shape, document.present))}>
          <Icon {...icon} /><span>{label}</span>
        </Menu.Item>
      ))}
    </MenuTool>
  )
}

const pickerRows = 8
const pickerColumns = 8

/** Pick a table's size by pointing at a grid, as in PowerPoint and Keynote. */
function TableTool({ documentId }: { readonly documentId: string }) {
  const [size, setSize] = useState({ rows: 3, columns: 3 })
  return (
    <Popover.Root>
      <Tip label="Table">
        <Popover.Trigger className="tool" aria-label="Table">
          <Table {...icon} />
          <ChevronDown {...tinyIcon} />
        </Popover.Trigger>
      </Tip>
      <Popover.Portal>
        <Popover.Positioner sideOffset={6} align="start">
          <Popover.Popup className="popover" finalFocus={false}>
            <div className="popover-label">Table: {size.rows} × {size.columns}</div>
            <div className="slides-table-picker" style={{ gridTemplateColumns: `repeat(${pickerColumns}, 16px)` }}>
              {Array.from({ length: pickerRows * pickerColumns }, (_, index) => {
                const row = Math.floor(index / pickerColumns) + 1
                const column = (index % pickerColumns) + 1
                return (
                  <Popover.Close
                    key={index}
                    className={`slides-table-cell${row <= size.rows && column <= size.columns ? ' is-chosen' : ''}`}
                    aria-label={`${row} rows by ${column} columns`}
                    onPointerEnter={() => setSize({ rows: row, columns: column })}
                    onFocus={() => setSize({ rows: row, columns: column })}
                    onClick={() => insertTable(documentId, row, column)}
                  />
                )
              })}
            </div>
          </Popover.Popup>
        </Popover.Positioner>
      </Popover.Portal>
    </Popover.Root>
  )
}

function ArrangeMenu({ documentId, disabled }: { readonly documentId: string; readonly disabled: boolean }) {
  return (
    <MenuTool icon={Layers} label="Arrange and align" disabled={disabled}>
      <Menu.Item className="menu-item" onClick={() => arrangeSelection(documentId, 'front')}><BringToFront {...icon} /><span>Bring to Front</span></Menu.Item>
      <Menu.Item className="menu-item" onClick={() => arrangeSelection(documentId, 'back')}><SendToBack {...icon} /><span>Send to Back</span></Menu.Item>
      <Menu.Separator className="menu-separator" />
      {objectAlignments.map(({ edge, label, icon: Icon }) => (
        <Menu.Item key={edge} className="menu-item" onClick={() => alignSelection(documentId, edge)}>
          <Icon {...icon} /><span>{label}</span>
        </Menu.Item>
      ))}
    </MenuTool>
  )
}

function TextBoxButton({ documentId, document }: ToolProps) {
  return <ToolButton icon={Type} label="Text box" onClick={() => addElement(documentId, newTextBox(document.present), true)} />
}

function PlayButton() {
  return <ToolButton icon={Play} label="Play slideshow" text="Play" command="slide.play" onClick={() => run('slide.play')} />
}

export function SlidesToolbar({ documentId, document }: ToolProps) {
  const toolbarShown = useViewStore(state => state.toolbar)
  // The split colour buttons remember the last colour chosen, like Keynote and Word.
  const [textColor, setTextColor] = useState('#c9352b')
  const [highlight, setHighlight] = useState<string | undefined>('#fff176')
  const [fillColor, setFillColor] = useState<string | undefined>('#4a78c2')
  const selected = selectedElements(document)
  const texts = selected.filter(holdsText)
  const table = selected.length === 1 && selected[0]?.kind === 'table' ? selected[0] : undefined
  const runs: readonly TextRun[] = texts.flatMap(element => allRuns(element.paragraphs))
  const paragraphs = texts.flatMap(element => element.paragraphs)
  const cell = table === undefined ? undefined : table.rows[document.cell?.row ?? 0]?.cells[document.cell?.column ?? 0]
  const every = (style: 'bold' | 'italic' | 'underline'): boolean => runs.length > 0 && runs.every(textRun => textRun[style])
  const noText = texts.length === 0 && table === undefined
  const noFill = !selected.some(element => element.kind !== 'image' && !(element.kind === 'shape' && (element.geometry.type === 'line' || element.geometry.type === 'arrow')))
  const font = runs[0]?.font ?? table?.font
  const size = runs[0]?.size ?? table?.size
  const align = paragraphs[0]?.align ?? cell?.align

  return (
    <>
      {toolbarShown && (
        <Toolbar label="Presentation tools">
          <ToolButton icon={Undo2} label="Undo" command="edit.undo" disabled={document.past.length === 0} onClick={() => run('edit.undo')} />
          <ToolButton icon={Redo2} label="Redo" command="edit.redo" disabled={document.future.length === 0} onClick={() => run('edit.redo')} />
          <ToolSeparator />
          <NewSlideMenu documentId={documentId} />
          <ToolSeparator />
          <TextBoxButton documentId={documentId} document={document} />
          <ShapeMenu documentId={documentId} document={document} />
          <ToolButton icon={ImagePlus} label="Picture…" onClick={() => void insertPicture(documentId).catch(reportFailure('Could not add the picture'))} />
          <TableTool documentId={documentId} />
          <ToolSeparator />
          <MenuTool label="Font" text={font ?? 'Font'} disabled={noText}>
            {fontChoices(font).map(name => (
              <Menu.Item key={name} className="menu-item" onClick={() => applyRunStyle(documentId, { font: name })}>
                <span style={{ fontFamily: `"${name}"` }}>{name}</span>
              </Menu.Item>
            ))}
          </MenuTool>
          <MenuTool label="Font size" text={size === undefined ? '–' : String(Math.round(size))} disabled={noText}>
            {fontSizes.map(value => (
              <Menu.Item key={value} className="menu-item" onClick={() => applyRunStyle(documentId, { size: value })}><span>{value} pt</span></Menu.Item>
            ))}
          </MenuTool>
          <ToolButton icon={Bold} label="Bold" command="format.bold" pressed={every('bold') || (runs.length === 0 && cell?.bold === true)} disabled={noText} onClick={() => toggleRunStyle(documentId, 'bold')} />
          <ToolButton icon={Italic} label="Italic" command="format.italic" pressed={every('italic')} disabled={texts.length === 0} onClick={() => toggleRunStyle(documentId, 'italic')} />
          <ToolButton icon={Underline} label="Underline" command="format.underline" pressed={every('underline')} disabled={texts.length === 0} onClick={() => toggleRunStyle(documentId, 'underline')} />
          <ColorTool
            icon={Baseline}
            label="Text colour"
            options={textColors}
            current={textColor}
            disabled={noText}
            onApply={value => { if (value !== undefined) applyRunStyle(documentId, { color: value }) }}
            onChoose={value => {
              if (value === undefined) return
              setTextColor(value)
              applyRunStyle(documentId, { color: value })
            }}
          />
          <ColorTool
            icon={Highlighter}
            label="Highlight"
            options={highlightColors}
            current={highlight}
            disabled={texts.length === 0}
            onApply={value => applyRunStyle(documentId, { highlight: value })}
            onChoose={value => {
              setHighlight(value)
              applyRunStyle(documentId, { highlight: value })
            }}
          />
          <ToolSeparator />
          <ToolButton icon={List} label="Bullets" pressed={paragraphs.length > 0 && paragraphs.every(paragraph => paragraph.list === 'bullet')} disabled={texts.length === 0} onClick={() => toggleList(documentId, 'bullet')} />
          <ToolButton icon={ListOrdered} label="Numbering" pressed={paragraphs.length > 0 && paragraphs.every(paragraph => paragraph.list === 'number')} disabled={texts.length === 0} onClick={() => toggleList(documentId, 'number')} />
          {alignments.map(alignment => (
            <ToolButton
              key={alignment.value}
              icon={alignment.icon}
              label={alignment.label}
              pressed={align === alignment.value}
              disabled={noText}
              onClick={() => setAlign(documentId, alignment.value)}
            />
          ))}
          <ToolSeparator />
          <ColorTool
            icon={PaintBucket}
            label="Fill colour"
            options={fillColors}
            current={fillColor}
            disabled={noFill}
            onApply={value => setFill(documentId, value)}
            onChoose={value => {
              setFillColor(value)
              setFill(documentId, value)
            }}
          />
          <ArrangeMenu documentId={documentId} disabled={selected.length === 0} />
          <span className="toolbar-grow" />
          <PlayButton />
        </Toolbar>
      )}

      <TitleEssentials>
        <NewSlideMenu documentId={documentId} />
        <TextBoxButton documentId={documentId} document={document} />
        <ShapeMenu documentId={documentId} document={document} />
        <PlayButton />
      </TitleEssentials>
    </>
  )
}
