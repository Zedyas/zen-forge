import {
  BetweenHorizontalStart,
  BetweenVerticalStart,
  BringToFront,
  Copy,
  FolderSearch,
  ImageMinus,
  ImageUp,
  IndentDecrease,
  IndentIncrease,
  List,
  ListOrdered,
  Minus,
  Plus,
  SendToBack,
  Trash2,
} from 'lucide-react'
import { useState } from 'react'
import { toast } from 'sonner'
import type { OpenDocument } from '../../app/documents-store'
import { platformClient } from '../../services/platform/client'
import { Inspector, InspectorRow, InspectorSection } from '../../ui/Inspector'
import { Tip } from '../../ui/Tip'
import { ToolButton } from '../../ui/Toolbar'
import {
  applyRunStyle,
  changeIndent,
  editTable,
  setBorder,
  setFill,
  setHeaderRow,
  setLineSpacing,
  setTableBorderColor,
  setVerticalAlign,
  toggleList,
} from './format-actions'
import {
  allRuns,
  canRotate,
  hasHeaderRow,
  holdsText,
  isLine,
  lineWidth,
  type SlideElement,
  type TableElement,
  type VerticalAlign,
} from './model'
import { backgroundColors, fillColors, fontChoices, lineSpacings, textColors } from './palette'
import {
  alignSelection,
  arrangeSelection,
  changeElement,
  chooseBackgroundPicture,
  duplicateSelection,
  removeSelection,
  setBackground,
  setTheme,
} from './slides-actions'
import { currentSlide, selectedElements, type ReadySlides } from './slides-store'
import { objectAlignments } from './SlidesToolbar'
import { themes } from './themes'

function points(value: number): string {
  return `${Math.round(value)} pt`
}

function reportFailure(title: string) {
  return (error: unknown): void => {
    toast.error(title, { description: error instanceof Error ? error.message : undefined })
  }
}

/** A number typed in the inspector; applied on Return or when the field loses focus, reverted by Escape. */
function NumberField({ label, value, suffix, onCommit }: { readonly label: string; readonly value: number; readonly suffix: string; onCommit(value: number): void }) {
  const [draft, setDraft] = useState<string>()
  const commit = (): void => {
    const parsed = Number.parseFloat(draft ?? '')
    setDraft(undefined)
    if (Number.isFinite(parsed) && Math.abs(parsed - value) > 0.001) onCommit(parsed)
  }
  return (
    <label className="slides-number">
      <span>{label}</span>
      <input
        className="text-field"
        inputMode="decimal"
        value={draft ?? String(Math.round(value * 10) / 10)}
        aria-label={label}
        onChange={event => setDraft(event.currentTarget.value)}
        onBlur={commit}
        onKeyDown={event => {
          if (event.key === 'Enter') event.currentTarget.blur()
          if (event.key === 'Escape') {
            setDraft(undefined)
            event.currentTarget.blur()
          }
        }}
      />
      <small>{suffix}</small>
    </label>
  )
}

/** A labelled row of colour swatches, stacked so eight fit across the inspector. */
function Swatches({ label, options, current, onChoose }: {
  readonly label: string
  readonly options: ReadonlyArray<{ readonly label: string; readonly value: string | undefined }>
  readonly current: string | undefined
  onChoose(value: string | undefined): void
}) {
  return (
    <div className="slides-field">
      <span>{label}</span>
      <div className="slides-swatches" role="group" aria-label={label}>
        {options.map(option => (
          <Tip key={option.label} label={option.label}>
            <button
              type="button"
              className={`swatch${option.value === undefined ? ' is-none' : ''}`}
              style={option.value === undefined ? undefined : { background: option.value }}
              aria-label={option.label}
              aria-pressed={option.value?.toLowerCase() === current?.toLowerCase()}
              onClick={() => onChoose(option.value)}
            />
          </Tip>
        ))}
      </div>
    </div>
  )
}

function Stepper({ label, value, step, minimum, onChange }: { readonly label: string; readonly value: number; readonly step: number; readonly minimum: number; onChange(value: number): void }) {
  return (
    <span className="slides-stepper">
      <Tip label={`Less ${label.toLowerCase()}`}>
        <button type="button" className="tool" aria-label={`Less ${label.toLowerCase()}`} onClick={() => onChange(Math.max(minimum, value - step))}><Minus aria-hidden="true" size={13} /></button>
      </Tip>
      <span>{Math.round(value * 10) / 10} pt</span>
      <Tip label={`More ${label.toLowerCase()}`}>
        <button type="button" className="tool" aria-label={`More ${label.toLowerCase()}`} onClick={() => onChange(value + step)}><Plus aria-hidden="true" size={13} /></button>
      </Tip>
    </span>
  )
}

function describe(element: SlideElement): string {
  switch (element.kind) {
    case 'text':
      return 'Text box'
    case 'image':
      return 'Picture'
    case 'table':
      return `Table, ${element.rows.length} × ${element.columns.length}`
    case 'shape':
      return element.geometry.type === 'line' ? 'Line' : element.geometry.type === 'arrow' ? 'Line with arrowhead' : 'Shape'
  }
}

function Placement({ documentId, element }: { readonly documentId: string; readonly element: SlideElement }) {
  const change = (patch: Partial<Pick<SlideElement, 'x' | 'y' | 'width' | 'height' | 'rotation'>>): void =>
    changeElement(documentId, element.id, current => ({ ...current, ...patch }))
  const sizeable = element.kind !== 'table'
  return (
    <InspectorSection title="Position and size">
      <div className="slides-number-grid">
        <NumberField label="X" value={element.x} suffix="pt" onCommit={x => change({ x })} />
        <NumberField label="Y" value={element.y} suffix="pt" onCommit={y => change({ y })} />
        {sizeable && <NumberField label="W" value={element.width} suffix="pt" onCommit={width => change({ width: Math.max(isLine(element) ? 0 : 1, width) })} />}
        {sizeable && <NumberField label="H" value={element.height} suffix="pt" onCommit={height => change({ height: Math.max(isLine(element) ? 0 : 1, height) })} />}
        {canRotate(element) && <NumberField label="Angle" value={element.rotation} suffix="°" onCommit={rotation => change({ rotation: ((rotation % 360) + 360) % 360 })} />}
      </div>
    </InspectorSection>
  )
}

function FillAndBorder({ documentId, element }: { readonly documentId: string; readonly element: SlideElement }) {
  if (element.kind !== 'text' && element.kind !== 'shape') return null
  const line = isLine(element)
  const border = element.border
  return (
    <InspectorSection title={line ? 'Line' : 'Fill and border'}>
      {!line && <Swatches label="Fill" options={fillColors} current={element.fill} onChoose={fill => setFill(documentId, fill)} />}
      <Swatches
        label={line ? 'Colour' : 'Border'}
        options={line ? textColors : [{ label: 'No border', value: undefined }, ...textColors.slice(0, 7)]}
        current={border?.color}
        onChoose={color => setBorder(documentId, color === undefined ? undefined : { color, width: border?.width ?? lineWidth })}
      />
      {border !== undefined && (
        <InspectorRow label="Width">
          <Stepper label="Width" value={border.width} step={0.5} minimum={0.5} onChange={width => setBorder(documentId, { ...border, width })} />
        </InspectorRow>
      )}
    </InspectorSection>
  )
}

const verticalAligns: ReadonlyArray<readonly [VerticalAlign, string]> = [['top', 'Top'], ['middle', 'Middle'], ['bottom', 'Bottom']]

function TextOptions({ documentId, elements }: { readonly documentId: string; readonly elements: readonly SlideElement[] }) {
  const texts = elements.filter(holdsText)
  const first = texts[0]
  if (first === undefined) return null
  const runs = allRuns(first.paragraphs)
  const font = runs[0]?.font ?? 'Arial'
  const size = runs[0]?.size ?? 18
  const paragraphs = texts.flatMap(element => element.paragraphs)
  const spacing = paragraphs[0]?.lineSpacing ?? 1
  const listed = (list: 'bullet' | 'number'): boolean => paragraphs.length > 0 && paragraphs.every(paragraph => paragraph.list === list)
  return (
    <InspectorSection title="Text">
      <InspectorRow label="Font">
        <select className="tool-select" value={font} aria-label="Font" onChange={event => applyRunStyle(documentId, { font: event.currentTarget.value })}>
          {fontChoices(font).map(name => <option key={name} value={name}>{name}</option>)}
        </select>
      </InspectorRow>
      <InspectorRow label="Size">
        <Stepper label="Size" value={size} step={2} minimum={4} onChange={value => applyRunStyle(documentId, { size: value })} />
      </InspectorRow>
      <InspectorRow label="Line spacing">
        <select className="tool-select" value={String(spacing)} aria-label="Line spacing" onChange={event => setLineSpacing(documentId, Number(event.currentTarget.value))}>
          {(lineSpacings.includes(spacing) ? lineSpacings : [spacing, ...lineSpacings]).map(value => <option key={value} value={String(value)}>{value.toFixed(value % 1 === 0 ? 1 : 2)}</option>)}
        </select>
      </InspectorRow>
      <InspectorRow label="Vertical">
        <span className="segmented">
          {verticalAligns.map(([value, label]) => (
            <button key={value} type="button" aria-pressed={first.verticalAlign === value} onClick={() => setVerticalAlign(documentId, value)}>{label}</button>
          ))}
        </span>
      </InspectorRow>
      <InspectorRow label="Lists">
        <span className="slides-inline-tools">
          <ToolButton icon={List} label="Bullets" pressed={listed('bullet')} onClick={() => toggleList(documentId, 'bullet')} />
          <ToolButton icon={ListOrdered} label="Numbering" pressed={listed('number')} onClick={() => toggleList(documentId, 'number')} />
          <ToolButton icon={IndentDecrease} label="Decrease indent" shortcut="⇧⇥" onClick={() => changeIndent(documentId, -1)} />
          <ToolButton icon={IndentIncrease} label="Increase indent" shortcut="⇥" onClick={() => changeIndent(documentId, 1)} />
        </span>
      </InspectorRow>
    </InspectorSection>
  )
}

function TableOptions({ documentId, table, cell }: { readonly documentId: string; readonly table: TableElement; readonly cell: ReadySlides['cell'] }) {
  const current = cell === undefined ? undefined : table.rows[cell.row]?.cells[cell.column]
  return (
    <InspectorSection title="Table">
      <p className="inspector-note">{cell === undefined ? 'Click a cell to choose where rows and columns go. Double-click to type.' : `Row ${cell.row + 1}, column ${cell.column + 1}.`}</p>
      <div className="slides-table-actions">
        <button type="button" className="button is-quiet" onClick={() => editTable(documentId, 'rowAbove')}><BetweenHorizontalStart aria-hidden="true" size={14} />Row above</button>
        <button type="button" className="button is-quiet" onClick={() => editTable(documentId, 'rowBelow')}><BetweenHorizontalStart aria-hidden="true" size={14} />Row below</button>
        <button type="button" className="button is-quiet" onClick={() => editTable(documentId, 'columnLeft')}><BetweenVerticalStart aria-hidden="true" size={14} />Column left</button>
        <button type="button" className="button is-quiet" onClick={() => editTable(documentId, 'columnRight')}><BetweenVerticalStart aria-hidden="true" size={14} />Column right</button>
        <button type="button" className="button is-quiet" disabled={table.rows.length === 1} onClick={() => editTable(documentId, 'deleteRow')}><Trash2 aria-hidden="true" size={14} />Delete row</button>
        <button type="button" className="button is-quiet" disabled={table.columns.length === 1} onClick={() => editTable(documentId, 'deleteColumn')}><Trash2 aria-hidden="true" size={14} />Delete column</button>
      </div>
      <label className="slides-check-row">
        <input type="checkbox" checked={hasHeaderRow(table)} onChange={event => setHeaderRow(documentId, event.currentTarget.checked)} />
        Header row
      </label>
      <Swatches label={cell === undefined ? 'Fill for every cell' : 'Cell fill'} options={fillColors} current={current?.fill} onChoose={fill => setFill(documentId, fill)} />
      <Swatches label="Borders" options={textColors} current={table.border.width > 0 ? table.border.color : undefined} onChoose={color => { if (color !== undefined) setTableBorderColor(documentId, color) }} />
      <InspectorRow label="Font">
        <select className="tool-select" value={table.font} aria-label="Table font" onChange={event => applyRunStyle(documentId, { font: event.currentTarget.value })}>
          {fontChoices(table.font).map(name => <option key={name} value={name}>{name}</option>)}
        </select>
      </InspectorRow>
      <InspectorRow label="Size">
        <Stepper label="Size" value={table.size} step={1} minimum={6} onChange={value => applyRunStyle(documentId, { size: value })} />
      </InspectorRow>
    </InspectorSection>
  )
}

function Arrange({ documentId, count }: { readonly documentId: string; readonly count: number }) {
  return (
    <InspectorSection title="Arrange">
      <div className="slides-align-row" role="group" aria-label={count > 1 ? 'Align objects' : 'Align to slide'}>
        {objectAlignments.map(({ edge, label, icon }) => (
          <ToolButton key={edge} icon={icon} label={count > 1 ? label : `${label} to the slide`} onClick={() => alignSelection(documentId, edge)} />
        ))}
      </div>
      <div className="slides-actions">
        <button type="button" className="button is-quiet" onClick={() => arrangeSelection(documentId, 'front')}><BringToFront aria-hidden="true" size={14} />Bring to front</button>
        <button type="button" className="button is-quiet" onClick={() => arrangeSelection(documentId, 'back')}><SendToBack aria-hidden="true" size={14} />Send to back</button>
        <button type="button" className="button is-quiet" onClick={() => duplicateSelection(documentId)}><Copy aria-hidden="true" size={14} />Duplicate<kbd>⌘D</kbd></button>
        <button type="button" className="button is-quiet" onClick={() => removeSelection(documentId)}><Trash2 aria-hidden="true" size={14} />Delete<kbd>⌫</kbd></button>
      </div>
    </InspectorSection>
  )
}

function ratio(width: number, height: number): string {
  const value = width / height
  if (Math.abs(value - 16 / 9) < 0.01) return '16:9'
  if (Math.abs(value - 4 / 3) < 0.01) return '4:3'
  if (Math.abs(value - 16 / 10) < 0.01) return '16:10'
  return `${(width / 72).toFixed(2)} × ${(height / 72).toFixed(2)} in`
}

function SlideOptions({ documentId, document }: { readonly documentId: string; readonly document: ReadySlides }) {
  const slide = currentSlide(document)
  if (slide === undefined) return null
  const index = document.present.slides.findIndex(candidate => candidate.id === slide.id)
  const { background, backgroundImage } = slide
  return (
    <>
      <InspectorSection title={`Slide ${index + 1}`}>
        <Swatches label="Background" options={backgroundColors} current={backgroundImage === undefined ? background : undefined} onChoose={value => { if (value !== undefined) setBackground(documentId, { background: value, backgroundImage: undefined }) }} />
        <div className="slides-actions">
          <button type="button" className="button is-quiet" onClick={() => void chooseBackgroundPicture(documentId).catch(reportFailure('Could not use the picture'))}>
            <ImageUp aria-hidden="true" size={14} />{backgroundImage === undefined ? 'Background picture…' : 'Change background picture…'}
          </button>
          {backgroundImage !== undefined && (
            <button type="button" className="button is-quiet" onClick={() => setBackground(documentId, { background, backgroundImage: undefined })}>
              <ImageMinus aria-hidden="true" size={14} />Remove background picture
            </button>
          )}
          <button type="button" className="button is-quiet" onClick={() => setBackground(documentId, { background, backgroundImage }, true)}>
            <Copy aria-hidden="true" size={14} />Use on every slide
          </button>
        </div>
      </InspectorSection>
      <InspectorSection title="Theme">
        <div className="slides-themes" role="group" aria-label="Theme">
          {themes.map(theme => (
            <Tip key={theme.id} label={`${theme.name}: recolours the slides and new objects`}>
              <button
                type="button"
                className="slides-theme"
                aria-label={theme.name}
                aria-pressed={document.present.theme === theme.id}
                style={{ background: theme.background, color: theme.text, fontFamily: `"${theme.font}"` }}
                onClick={() => setTheme(documentId, theme.id)}
              >
                Aa<i style={{ background: theme.accent }} />
              </button>
            </Tip>
          ))}
        </div>
      </InspectorSection>
    </>
  )
}

export function SlidesInspector({ documentId, document, openDocument }: {
  readonly documentId: string
  readonly document: ReadySlides
  readonly openDocument: OpenDocument
}) {
  const elements = selectedElements(document)
  const single = elements.length === 1 ? elements[0] : undefined
  const { width, height } = document.present
  return (
    <Inspector label="Presentation inspector">
      {single !== undefined && (
        <>
          <InspectorSection title={describe(single)}>
            <p className="inspector-note">Drag to move and handles to resize; Shift keeps proportions. Arrow keys nudge by 1 pt, with Shift by 10 pt.</p>
          </InspectorSection>
          <Placement documentId={documentId} element={single} />
          <FillAndBorder documentId={documentId} element={single} />
          {single.kind === 'table' && <TableOptions documentId={documentId} table={single} cell={document.cell} />}
        </>
      )}
      {elements.length > 1 && (
        <InspectorSection title={`${elements.length} objects selected`}>
          <p className="inspector-note">Shift-click adds or removes an object. They move, align and delete together.</p>
        </InspectorSection>
      )}
      {elements.length > 0 && <TextOptions documentId={documentId} elements={elements} />}
      {elements.length > 0 && <Arrange documentId={documentId} count={elements.length} />}
      {elements.length === 0 && <SlideOptions documentId={documentId} document={document} />}
      <InspectorSection title="Presentation">
        <InspectorRow label="Slides"><span>{document.present.slides.length}</span></InspectorRow>
        <InspectorRow label="Size"><span>{ratio(width, height)}, {points(width)} × {points(height)}</span></InspectorRow>
        <InspectorRow label="File"><span className="slides-file-name">{openDocument.name}.{openDocument.extension}</span></InspectorRow>
        {openDocument.path !== undefined && (
          <button type="button" className="button is-quiet slides-inspector-action" onClick={() => {
            if (openDocument.path !== undefined) void platformClient.revealInFinder(openDocument.path)
          }}>
            <FolderSearch aria-hidden="true" size={14} />Show in Finder
          </button>
        )}
      </InspectorSection>
    </Inspector>
  )
}
