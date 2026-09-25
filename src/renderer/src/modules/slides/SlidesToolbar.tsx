import { Menu } from '@base-ui/react/menu'
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Baseline,
  Bold,
  ChevronDown,
  Circle,
  ImagePlus,
  Italic,
  Minus,
  MoveRight,
  PaintBucket,
  Plus,
  Redo2,
  Shapes,
  Square,
  SquareRoundCorner,
  Type,
  Underline,
  Undo2,
  type LucideIcon,
} from 'lucide-react'
import { useState } from 'react'
import { toast } from 'sonner'
import type { CommandId } from '@shared/commands'
import { commandShortcut, Tip } from '../../ui/Tip'
import { ColorTool, Toolbar, ToolButton, ToolSeparator } from '../../ui/Toolbar'
import { TitleEssentials } from '../../ui/TitleSlot'
import { useViewStore } from '../../app/view-store'
import { allRuns, holdsText, newShape, newTextBox, slideLayouts, type BasicShape, type HorizontalAlign } from './model'
import {
  addElement,
  addSlide,
  applyRunStyle,
  insertPicture,
  runSlidesCommand,
  setAlign,
  setFill,
  toggleRunStyle,
} from './slides-actions'
import { fillColors, textColors } from './palette'
import { selectedElement, type ReadySlides } from './slides-store'

const fontSizes: readonly number[] = [10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 72, 96]

const shapes: ReadonlyArray<{ readonly shape: BasicShape; readonly label: string; readonly icon: LucideIcon }> = [
  { shape: 'rect', label: 'Rectangle', icon: Square },
  { shape: 'roundRect', label: 'Rounded rectangle', icon: SquareRoundCorner },
  { shape: 'ellipse', label: 'Ellipse', icon: Circle },
  { shape: 'line', label: 'Line', icon: Minus },
  { shape: 'arrow', label: 'Arrow', icon: MoveRight },
]

const alignments: ReadonlyArray<{ readonly value: HorizontalAlign; readonly label: string; readonly icon: LucideIcon }> = [
  { value: 'left', label: 'Align left', icon: AlignLeft },
  { value: 'center', label: 'Align centre', icon: AlignCenter },
  { value: 'right', label: 'Align right', icon: AlignRight },
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

function NewSlideMenu({ documentId }: { readonly documentId: string }) {
  return (
    <Menu.Root>
      <Tip label="New slide" shortcut={commandShortcut('slide.new')}>
        <Menu.Trigger className="tool" aria-label="New slide">
          <Plus aria-hidden="true" size={16} strokeWidth={1.7} />
          <ChevronDown aria-hidden="true" size={11} strokeWidth={2.2} />
        </Menu.Trigger>
      </Tip>
      <Menu.Portal>
        <Menu.Positioner sideOffset={6} align="start">
          <Menu.Popup className="menu-popup">
            {slideLayouts.map(layout => (
              <Menu.Item key={layout.id} className="menu-item" onClick={() => addSlide(documentId, layout.id)}>
                <span>{layout.label}</span>
              </Menu.Item>
            ))}
          </Menu.Popup>
        </Menu.Positioner>
      </Menu.Portal>
    </Menu.Root>
  )
}

function ShapeMenu({ documentId, document }: ToolProps) {
  return (
    <Menu.Root>
      <Tip label="Shape">
        <Menu.Trigger className="tool" aria-label="Shape">
          <Shapes aria-hidden="true" size={16} strokeWidth={1.7} />
          <ChevronDown aria-hidden="true" size={11} strokeWidth={2.2} />
        </Menu.Trigger>
      </Tip>
      <Menu.Portal>
        <Menu.Positioner sideOffset={6} align="start">
          <Menu.Popup className="menu-popup">
            {shapes.map(({ shape, label, icon: Icon }) => (
              <Menu.Item key={shape} className="menu-item" onClick={() => addElement(documentId, newShape(shape, document.present))}>
                <Icon aria-hidden="true" size={16} strokeWidth={1.7} /><span>{label}</span>
              </Menu.Item>
            ))}
          </Menu.Popup>
        </Menu.Positioner>
      </Menu.Portal>
    </Menu.Root>
  )
}

function TextBoxButton({ documentId, document }: ToolProps) {
  return <ToolButton icon={Type} label="Text box" onClick={() => addElement(documentId, newTextBox(document.present), true)} />
}

function FontSizeMenu({ documentId, size, disabled }: { readonly documentId: string; readonly size: number | undefined; readonly disabled: boolean }) {
  return (
    <Menu.Root>
      <Tip label="Font size">
        <Menu.Trigger className="tool is-label slides-size" aria-label="Font size" disabled={disabled}>
          <span>{size === undefined ? '–' : Math.round(size)}</span>
          <ChevronDown aria-hidden="true" size={11} strokeWidth={2.2} />
        </Menu.Trigger>
      </Tip>
      <Menu.Portal>
        <Menu.Positioner sideOffset={6} align="start">
          <Menu.Popup className="menu-popup slides-size-menu">
            {fontSizes.map(value => (
              <Menu.Item key={value} className="menu-item" onClick={() => applyRunStyle(documentId, { size: value })}>
                <span>{value} pt</span>
              </Menu.Item>
            ))}
          </Menu.Popup>
        </Menu.Positioner>
      </Menu.Portal>
    </Menu.Root>
  )
}

export function SlidesToolbar({ documentId, document }: ToolProps) {
  const toolbarShown = useViewStore(state => state.toolbar)
  // The split colour buttons remember the last colour chosen, like Keynote and Word.
  const [textColor, setTextColor] = useState('#c9352b')
  const [fillColor, setFillColor] = useState<string | undefined>('#4a78c2')
  const selected = selectedElement(document)
  const text = selected !== undefined && holdsText(selected) ? selected : undefined
  const runs = text === undefined ? [] : allRuns(text.paragraphs)
  const every = (style: 'bold' | 'italic' | 'underline'): boolean => runs.length > 0 && runs.every(textRun => textRun[style])
  const align = text?.paragraphs[0]?.align
  const noText = text === undefined
  const noFill = selected === undefined || selected.kind === 'image'

  return (
    <>
      {toolbarShown && (
        <Toolbar label="Presentation tools">
          <ToolButton icon={Undo2} label="Undo" command="edit.undo" disabled={document.past.length === 0} onClick={() => run('edit.undo')} />
          <ToolButton icon={Redo2} label="Redo" command="edit.redo" disabled={document.future.length === 0} onClick={() => run('edit.redo')} />
          <ToolSeparator />
          <NewSlideMenu documentId={documentId} />
          <TextBoxButton documentId={documentId} document={document} />
          <ShapeMenu documentId={documentId} document={document} />
          <ToolButton icon={ImagePlus} label="Picture…" onClick={() => void insertPicture(documentId).catch(reportFailure('Could not add the picture'))} />
          <ToolSeparator />
          <FontSizeMenu documentId={documentId} size={runs[0]?.size} disabled={noText} />
          <ToolButton icon={Bold} label="Bold" command="format.bold" pressed={every('bold')} disabled={noText} onClick={() => toggleRunStyle(documentId, 'bold')} />
          <ToolButton icon={Italic} label="Italic" command="format.italic" pressed={every('italic')} disabled={noText} onClick={() => toggleRunStyle(documentId, 'italic')} />
          <ToolButton icon={Underline} label="Underline" command="format.underline" pressed={every('underline')} disabled={noText} onClick={() => toggleRunStyle(documentId, 'underline')} />
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
          <ToolSeparator />
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
        </Toolbar>
      )}

      <TitleEssentials>
        <NewSlideMenu documentId={documentId} />
        <TextBoxButton documentId={documentId} document={document} />
        <ShapeMenu documentId={documentId} document={document} />
      </TitleEssentials>
    </>
  )
}
