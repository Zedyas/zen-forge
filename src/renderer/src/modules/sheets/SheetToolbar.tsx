import { Menu } from '@base-ui/react/menu'
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDownAZ,
  ArrowUpZA,
  Baseline,
  Bold,
  Braces,
  Calendar,
  Check,
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
  Strikethrough,
  TableCellsMerge,
  TableCellsSplit,
  Underline,
  Undo2,
  WrapText,
} from 'lucide-react'
import { useState, type ReactNode } from 'react'
import { CellPopover } from './CellPopover'
import { numberFormatPattern } from './io/number-format'
import { formatCellValue } from './model/format'
import { formatWithCode } from './model/format-code'
import { defaultCellStyle, defaultFontSize, type HorizontalAlign } from './model/workbook-data'
import type { NumberFormat } from './model/Workbook'
import {
  autoSum,
  fillColors,
  fontFamilies,
  fontSizes,
  mergeSelection,
  setFormatCode,
  setStyle,
  sortByActiveColumn,
  textColors,
  toggleStyle,
  type SheetController,
} from './sheet-commands'
import { icon, smallIcon, tinyIcon } from '../../ui/icons'
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
          <Icon {...icon} />
          <ChevronDown {...tinyIcon} />
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

/** A toolbar menu that names its current choice, like Sumi's font menus. */
function ChoiceMenu({ label, current, children }: { readonly label: string; readonly current: string; readonly children: ReactNode }) {
  return (
    <Menu.Root>
      <Tip label={label}>
        <Menu.Trigger className="tool is-label menu-label" aria-label={label}>
          <span>{current}</span>
          <ChevronDown {...tinyIcon} />
        </Menu.Trigger>
      </Tip>
      <Menu.Portal>
        <Menu.Positioner sideOffset={6} align="start">
          <Menu.Popup className="menu-popup sheet-choice-menu">{children}</Menu.Popup>
        </Menu.Positioner>
      </Menu.Portal>
    </Menu.Root>
  )
}

function Choice({ active, family, onClick, children }: { readonly active: boolean; readonly family?: string; readonly onClick: () => void; readonly children: ReactNode }) {
  return (
    <Menu.Item className="menu-item" data-active={active ? '' : undefined} onClick={onClick}>
      <span style={family === undefined ? undefined : { fontFamily: `"${family}"` }}>{children}</span>
      {active && <Check {...smallIcon} />}
    </Menu.Item>
  )
}

/** Font family and size. Excel's default size (11) and the default font are no formatting at all. */
function FontMenus({ controller }: { readonly controller: SheetController }) {
  const style = controller.workbook.getStyle(controller.sheetId, controller.cell)
  const families: readonly string[] = fontFamilies
  const extra = style.fontFamily !== undefined && !families.includes(style.fontFamily) ? [style.fontFamily] : []
  const size = style.fontSize ?? defaultFontSize
  return (
    <>
      <ChoiceMenu label="Font" current={style.fontFamily ?? 'Default'}>
        <Choice active={style.fontFamily === undefined} onClick={() => setStyle(controller, { fontFamily: undefined })}>Default</Choice>
        {[...families, ...extra].map(family => (
          <Choice key={family} active={style.fontFamily === family} family={family} onClick={() => setStyle(controller, { fontFamily: family })}>{family}</Choice>
        ))}
      </ChoiceMenu>
      <ChoiceMenu label="Font size" current={String(size)}>
        {[...new Set([...fontSizes, size])].sort((left, right) => left - right).map(value => (
          <Choice key={value} active={value === size} onClick={() => setStyle(controller, { fontSize: value === defaultFontSize ? undefined : value })}>{value}</Choice>
        ))}
      </ChoiceMenu>
    </>
  )
}

const formatCodeExamples = ['#,##0.00;[Red]-#,##0.00', '0.0%', '€#,##0.00', '0.00E+00', '0.0,,"M"', 'd mmm yyyy', 'dddd', 'h:mm AM/PM']

/** "Custom…" number format: an Excel format code, previewed on the active cell's value. */
function FormatCodeTool({ controller }: { readonly controller: SheetController }) {
  const { workbook, sheetId, cell } = controller
  const style = workbook.getStyle(sheetId, cell)
  const value = workbook.getCellValue(sheetId, cell)
  const sample = value === null || value === '' ? 1234.5 : value
  return (
    <CellPopover
      icon={Braces}
      label="Custom number format"
      pressed={style.formatCode !== undefined}
      initial={style.formatCode ?? numberFormatPattern(style) ?? 'General'}
      placeholder="Excel format code, like #,##0.00"
      submitLabel="Apply"
      onSubmit={code => {
        setFormatCode(controller, code)
        return true
      }}
    >
      {(draft, setDraft) => (
        <>
          <p className="sheet-cell-popover-note">
            {draft.trim() === '' || draft.trim().toLowerCase() === 'general'
              ? formatCellValue(sample, defaultCellStyle)
              : formatWithCode(sample, draft)?.text ?? 'Ledger shows this code as General, and keeps it in the file.'}
          </p>
          <div className="sheet-format-examples">
            {formatCodeExamples.map(example => (
              <button key={example} type="button" className="button is-quiet" onClick={() => setDraft(example)}><code>{example}</code></button>
            ))}
          </div>
        </>
      )}
    </CellPopover>
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
        <FontMenus controller={controller} />
        <ToolButton icon={Bold} label="Bold" command="format.bold" pressed={style.bold === true} onClick={() => toggleStyle(controller, 'bold')} />
        <ToolButton icon={Italic} label="Italic" command="format.italic" pressed={style.italic === true} onClick={() => toggleStyle(controller, 'italic')} />
        <ToolButton icon={Underline} label="Underline" command="format.underline" pressed={style.underline === true} onClick={() => toggleStyle(controller, 'underline')} />
        <ToolButton icon={Strikethrough} label="Strikethrough" command="format.strikethrough" pressed={style.strike === true} onClick={() => toggleStyle(controller, 'strike')} />
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
        <ToolButton icon={WrapText} label="Wrap text" command="format.wrapText" pressed={style.wrap === true} onClick={() => toggleStyle(controller, 'wrap')} />
        {workbook.mergeAt(sheetId, cell) === undefined
          ? <ToolButton icon={TableCellsMerge} label="Merge cells" command="format.merge" onClick={() => void mergeSelection(controller, false)} />
          : <ToolButton icon={TableCellsSplit} label="Unmerge cells" command="format.unmerge" onClick={() => workbook.unmergeCells(sheetId, controller.range)} />}
        <ToolSeparator />
        {numberFormats.map(format => (
          <ToolButton
            key={format.value}
            icon={format.icon}
            text={format.text}
            label={format.label}
            pressed={style.numberFormat === format.value && style.formatCode === undefined}
            onClick={() => setStyle(controller, { numberFormat: format.value })}
          />
        ))}
        <FormatCodeTool controller={controller} />
        <ToolButton icon={DecimalsArrowLeft} label="Fewer decimal places" onClick={() => workbook.changeDecimalPlaces(sheetId, controller.areas, -1)} />
        <ToolButton icon={DecimalsArrowRight} label="More decimal places" onClick={() => workbook.changeDecimalPlaces(sheetId, controller.areas, 1)} />
        <ToolSeparator />
        <ToolButton icon={Sigma} label="AutoSum" onClick={() => autoSum(controller)} />
        <ToolButton icon={ArrowDownAZ} label="Sort A to Z by this column" onClick={() => sortByActiveColumn(controller, true)} />
        <ToolButton icon={ArrowUpZA} label="Sort Z to A by this column" onClick={() => sortByActiveColumn(controller, false)} />
        <ToolSeparator />
        <StructureMenu controller={controller} axis="rows" />
        <StructureMenu controller={controller} axis="columns" />
        <span className="toolbar-grow" />
        <ToolButton icon={RemoveFormatting} label="Clear formatting" command="format.clear" onClick={() => workbook.clearFormatting(sheetId, controller.areas)} />
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
