import { Menu } from '@base-ui/react/menu'
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDownAZ,
  ArrowUpZA,
  Baseline,
  Bold,
  Calendar,
  ChevronDown,
  Columns3,
  DecimalsArrowLeft,
  DecimalsArrowRight,
  DollarSign,
  Hash,
  Italic,
  PaintBucket,
  Percent,
  Redo2,
  RemoveFormatting,
  Rows3,
  Search,
  Sigma,
  Underline,
  Undo2,
} from 'lucide-react'
import { useState } from 'react'
import type { HorizontalAlign } from './model/workbook-data'
import type { NumberFormat } from './model/Workbook'
import {
  autoSum,
  fillColors,
  setStyle,
  sortByActiveColumn,
  textColors,
  toggleStyle,
  type SheetController,
} from './sheet-commands'
import { Tip } from '../../ui/Tip'
import { ColorTool, Toolbar, ToolButton, ToolSeparator } from '../../ui/Toolbar'
import { TitleEssentials } from '../../ui/TitleSlot'
import { useViewStore } from '../../app/view-store'

const numberFormats: ReadonlyArray<{ readonly value: NumberFormat; readonly label: string; readonly icon?: typeof Hash; readonly text?: string }> = [
  { value: 'general', label: 'General', text: '123' },
  { value: 'number', label: 'Number with separators', icon: Hash },
  { value: 'currency', label: 'Currency', icon: DollarSign },
  { value: 'percent', label: 'Percent', icon: Percent },
  { value: 'date', label: 'Date', icon: Calendar },
]

const alignments: ReadonlyArray<{ readonly value: HorizontalAlign; readonly label: string; readonly icon: typeof AlignLeft }> = [
  { value: 'left', label: 'Align left', icon: AlignLeft },
  { value: 'center', label: 'Align centre', icon: AlignCenter },
  { value: 'right', label: 'Align right', icon: AlignRight },
]

interface StructureMenuProps {
  readonly controller: SheetController
  readonly axis: 'rows' | 'columns'
}

/** Row and column structure: rarer actions, so they sit behind one menu per axis. */
function StructureMenu({ controller, axis }: StructureMenuProps) {
  const { workbook, sheetId, range } = controller
  const isRows = axis === 'rows'
  const start = isRows ? range.y : range.x
  const count = isRows ? range.height : range.width
  const noun = `${count} ${isRows ? 'row' : 'column'}${count === 1 ? '' : 's'}`
  const items: ReadonlyArray<readonly [string, () => void]> = isRows
    ? [
        [`Insert ${noun} above`, () => workbook.insertRows(sheetId, start, count)],
        [`Insert ${noun} below`, () => workbook.insertRows(sheetId, start + count, count)],
        [`Delete ${noun}`, () => workbook.deleteRows(sheetId, start, count)],
        [`Hide ${noun}`, () => workbook.hideRows(sheetId, start, count)],
        ['Show all rows', () => workbook.showAllRows(sheetId)],
        [`Freeze through row ${start + count}`, () => workbook.freezeRows(sheetId, start + count)],
        ['Unfreeze rows', () => workbook.freezeRows(sheetId, 0)],
      ]
    : [
        [`Insert ${noun} left`, () => workbook.insertColumns(sheetId, start, count)],
        [`Insert ${noun} right`, () => workbook.insertColumns(sheetId, start + count, count)],
        [`Delete ${noun}`, () => workbook.deleteColumns(sheetId, start, count)],
        [`Hide ${noun}`, () => workbook.hideColumns(sheetId, start, count)],
        ['Show all columns', () => workbook.showAllColumns(sheetId)],
        [`Freeze through this column`, () => workbook.freezeColumns(sheetId, start + count)],
        ['Unfreeze columns', () => workbook.freezeColumns(sheetId, 0)],
      ]
  const Icon = isRows ? Rows3 : Columns3
  return (
    <Menu.Root>
      <Tip label={isRows ? 'Rows' : 'Columns'}>
        <Menu.Trigger className="tool" aria-label={isRows ? 'Rows' : 'Columns'}>
          <Icon aria-hidden="true" size={16} strokeWidth={1.7} />
          <ChevronDown aria-hidden="true" size={11} strokeWidth={2.2} />
        </Menu.Trigger>
      </Tip>
      <Menu.Portal>
        <Menu.Positioner sideOffset={6} align="start">
          <Menu.Popup className="menu-popup">
            {items.map(([label, action], index) => (
              <div key={label}>
                {(index === 3 || index === 5) && <Menu.Separator className="menu-separator" />}
                <Menu.Item className="menu-item" onClick={action}><span>{label}</span></Menu.Item>
              </div>
            ))}
          </Menu.Popup>
        </Menu.Positioner>
      </Menu.Portal>
    </Menu.Root>
  )
}

interface SheetToolbarProps {
  readonly controller: SheetController
  readonly findOpen: boolean
}

export function SheetToolbar({ controller, findOpen }: SheetToolbarProps) {
  const { workbook, sheetId, cell } = controller
  const style = workbook.getStyle(sheetId, cell)
  // The split colour buttons remember the last colour chosen in this window, like Numbers and Word.
  const [textColor, setTextColor] = useState<string | undefined>('#c9352b')
  const [fillColor, setFillColor] = useState<string | undefined>('#fdf1c7')
  const toolbarShown = useViewStore(state => state.toolbar)

  return (
    <>
      {toolbarShown && <Toolbar label="Spreadsheet tools">
        <ToolButton icon={Undo2} label="Undo" command="edit.undo" disabled={!workbook.canUndo()} onClick={() => workbook.undo()} />
        <ToolButton icon={Redo2} label="Redo" command="edit.redo" disabled={!workbook.canRedo()} onClick={() => workbook.redo()} />
        <ToolSeparator />
        <ToolButton icon={Bold} label="Bold" command="format.bold" pressed={style.bold === true} onClick={() => toggleStyle(controller, 'bold')} />
        <ToolButton icon={Italic} label="Italic" command="format.italic" pressed={style.italic === true} onClick={() => toggleStyle(controller, 'italic')} />
        <ToolButton icon={Underline} label="Underline" command="format.underline" pressed={style.underline === true} onClick={() => toggleStyle(controller, 'underline')} />
        <ColorTool
          icon={Baseline}
          label="Text colour"
          options={textColors}
          current={textColor}
          onApply={value => setStyle(controller, { textColor: value })}
          onChoose={value => {
            setTextColor(value)
            setStyle(controller, { textColor: value })
          }}
        />
        <ColorTool
          icon={PaintBucket}
          label="Fill colour"
          options={fillColors}
          current={fillColor}
          onApply={value => setStyle(controller, { fillColor: value })}
          onChoose={value => {
            setFillColor(value)
            setStyle(controller, { fillColor: value })
          }}
        />
        <ToolSeparator />
        {alignments.map(alignment => (
          <ToolButton
            key={alignment.value}
            icon={alignment.icon}
            label={alignment.label}
            pressed={style.align === alignment.value}
            onClick={() => setStyle(controller, { align: style.align === alignment.value ? undefined : alignment.value })}
          />
        ))}
        <ToolSeparator />
        {numberFormats.map(format => (
          <ToolButton
            key={format.value}
            icon={format.icon}
            text={format.text}
            label={format.label}
            pressed={style.numberFormat === format.value}
            onClick={() => setStyle(controller, { numberFormat: format.value })}
          />
        ))}
        <ToolButton icon={DecimalsArrowLeft} label="Fewer decimal places" onClick={() => workbook.changeDecimalPlaces(sheetId, controller.range, -1)} />
        <ToolButton icon={DecimalsArrowRight} label="More decimal places" onClick={() => workbook.changeDecimalPlaces(sheetId, controller.range, 1)} />
        <ToolSeparator />
        <ToolButton icon={Sigma} label="AutoSum" onClick={() => autoSum(controller)} />
        <ToolButton icon={ArrowDownAZ} label="Sort A to Z by this column" onClick={() => sortByActiveColumn(controller, true)} />
        <ToolButton icon={ArrowUpZA} label="Sort Z to A by this column" onClick={() => sortByActiveColumn(controller, false)} />
        <ToolSeparator />
        <StructureMenu controller={controller} axis="rows" />
        <StructureMenu controller={controller} axis="columns" />
        <span className="toolbar-grow" />
        <ToolButton icon={RemoveFormatting} label="Clear formatting" command="format.clear" onClick={() => workbook.clearFormatting(sheetId, controller.range)} />
        <ToolButton icon={Search} label="Find and replace" command="edit.find" pressed={findOpen} onClick={controller.openFind} />
      </Toolbar>}

      <TitleEssentials>
        <ToolButton icon={Undo2} label="Undo" command="edit.undo" disabled={!workbook.canUndo()} onClick={() => workbook.undo()} />
        <ToolButton icon={Bold} label="Bold" command="format.bold" pressed={style.bold === true} onClick={() => toggleStyle(controller, 'bold')} />
        <ColorTool
          icon={PaintBucket}
          label="Fill colour"
          options={fillColors}
          current={fillColor}
          onApply={value => setStyle(controller, { fillColor: value })}
          onChoose={value => {
            setFillColor(value)
            setStyle(controller, { fillColor: value })
          }}
        />
        <ToolButton icon={DollarSign} label="Currency" pressed={style.numberFormat === 'currency'} onClick={() => setStyle(controller, { numberFormat: 'currency' })} />
        <ToolButton icon={Sigma} label="AutoSum" onClick={() => autoSum(controller)} />
        <ToolButton icon={ArrowDownAZ} label="Sort A to Z by this column" onClick={() => sortByActiveColumn(controller, true)} />
      </TitleEssentials>
    </>
  )
}
