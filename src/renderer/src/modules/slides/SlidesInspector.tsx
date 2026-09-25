import { BringToFront, Copy, FolderSearch, IndentDecrease, IndentIncrease, List, Minus, Plus, SendToBack, Trash2 } from 'lucide-react'
import { useState } from 'react'
import type { OpenDocument } from '../../app/documents-store'
import { platformClient } from '../../services/platform/client'
import { Inspector, InspectorRow, InspectorSection } from '../../ui/Inspector'
import { Tip } from '../../ui/Tip'
import { ToolButton } from '../../ui/Toolbar'
import {
  allRuns,
  holdsText,
  isLine,
  lineWidth,
  type ShapeElement,
  type SlideElement,
  type TextBoxElement,
  type VerticalAlign,
} from './model'
import {
  applyRunStyle,
  arrangeSelection,
  changeElement,
  duplicateSelection,
  removeElement,
  setSlideBackground,
} from './slides-actions'
import { currentSlide, selectedElement, type ReadySlides } from './slides-store'
import { fillColors, textColors } from './palette'

const fonts: readonly string[] = ['Arial', 'Helvetica', 'Avenir Next', 'Gill Sans', 'Verdana', 'Trebuchet MS', 'Georgia', 'Times New Roman', 'Courier New']

const backgrounds: ReadonlyArray<{ readonly label: string; readonly value: string }> = [
  { label: 'White', value: '#ffffff' },
  { label: 'Paper', value: '#f7f4ec' },
  { label: 'Light grey', value: '#eef0f3' },
  { label: 'Light blue', value: '#e3ecfa' },
  { label: 'Light green', value: '#e2f4ea' },
  { label: 'Charcoal', value: '#2b2d31' },
  { label: 'Navy', value: '#1d2b4f' },
  { label: 'Black', value: '#000000' },
]

function points(value: number): string {
  return `${Math.round(value)} pt`
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

function Swatches({ label, options, current, onChoose }: {
  readonly label: string
  readonly options: ReadonlyArray<{ readonly label: string; readonly value: string | undefined }>
  readonly current: string | undefined
  onChoose(value: string | undefined): void
}) {
  return (
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
  if (element.kind === 'text') return 'Text box'
  if (element.kind === 'image') return 'Picture'
  if (element.geometry.type === 'line') return 'Line'
  if (element.geometry.type === 'arrow') return 'Arrow'
  return 'Shape'
}

function Placement({ documentId, element }: { readonly documentId: string; readonly element: SlideElement }) {
  const change = (patch: Partial<Pick<SlideElement, 'x' | 'y' | 'width' | 'height' | 'rotation'>>): void =>
    changeElement(documentId, element.id, current => ({ ...current, ...patch }))
  return (
    <InspectorSection title="Position and size">
      <div className="slides-number-grid">
        <NumberField label="X" value={element.x} suffix="pt" onCommit={x => change({ x })} />
        <NumberField label="Y" value={element.y} suffix="pt" onCommit={y => change({ y })} />
        <NumberField label="W" value={element.width} suffix="pt" onCommit={width => change({ width: Math.max(isLine(element) ? 0 : 1, width) })} />
        <NumberField label="H" value={element.height} suffix="pt" onCommit={height => change({ height: Math.max(isLine(element) ? 0 : 1, height) })} />
        {!isLine(element) && (
          <NumberField label="Rotation" value={element.rotation} suffix="°" onCommit={rotation => change({ rotation: ((rotation % 360) + 360) % 360 })} />
        )}
      </div>
    </InspectorSection>
  )
}

function FillAndBorder({ documentId, element }: { readonly documentId: string; readonly element: TextBoxElement | ShapeElement }) {
  const line = isLine(element)
  const change = (patch: Partial<Pick<ShapeElement, 'fill' | 'border'>>): void =>
    changeElement(documentId, element.id, current => current.kind === 'image' ? current : { ...current, ...patch })
  const border = element.border
  return (
    <InspectorSection title={line ? 'Line' : 'Fill and border'}>
      {!line && (
        <InspectorRow label="Fill">
          <Swatches label="Fill" options={fillColors} current={element.fill} onChoose={fill => change({ fill })} />
        </InspectorRow>
      )}
      <InspectorRow label={line ? 'Colour' : 'Border'}>
        <Swatches
          label={line ? 'Line colour' : 'Border colour'}
          options={line ? textColors : [{ label: 'No border', value: undefined }, ...textColors.slice(0, 7)]}
          current={border?.color}
          onChoose={color => change({ border: color === undefined ? undefined : { color, width: border?.width ?? lineWidth } })}
        />
      </InspectorRow>
      {border !== undefined && (
        <InspectorRow label="Width">
          <Stepper label="Width" value={border.width} step={0.5} minimum={0.5} onChange={width => change({ border: { ...border, width } })} />
        </InspectorRow>
      )}
    </InspectorSection>
  )
}

const verticalAligns: ReadonlyArray<readonly [VerticalAlign, string]> = [['top', 'Top'], ['middle', 'Middle'], ['bottom', 'Bottom']]

function TextOptions({ documentId, element }: { readonly documentId: string; readonly element: TextBoxElement | ShapeElement }) {
  const runs = allRuns(element.paragraphs)
  const font = runs[0]?.font ?? 'Arial'
  const size = runs[0]?.size ?? 18
  const bulleted = element.paragraphs.length > 0 && element.paragraphs.every(paragraph => paragraph.bullet)
  const change = (next: (current: TextBoxElement | ShapeElement) => TextBoxElement | ShapeElement): void =>
    changeElement(documentId, element.id, current => holdsText(current) ? next(current) : current)
  const indent = (step: number): void => change(current => ({
    ...current,
    paragraphs: current.paragraphs.map(paragraph => ({ ...paragraph, level: Math.min(8, Math.max(0, paragraph.level + step)) })),
  }))
  return (
    <InspectorSection title="Text">
      <InspectorRow label="Font">
        <select className="tool-select" value={font} aria-label="Font" onChange={event => applyRunStyle(documentId, { font: event.currentTarget.value })}>
          {(fonts.includes(font) ? fonts : [font, ...fonts]).map(name => <option key={name} value={name}>{name}</option>)}
        </select>
      </InspectorRow>
      <InspectorRow label="Size">
        <Stepper label="Size" value={size} step={2} minimum={4} onChange={value => applyRunStyle(documentId, { size: value })} />
      </InspectorRow>
      <InspectorRow label="Vertical">
        <span className="segmented">
          {verticalAligns.map(([value, label]) => (
            <button key={value} type="button" aria-pressed={element.verticalAlign === value} onClick={() => change(current => ({ ...current, verticalAlign: value }))}>
              {label}
            </button>
          ))}
        </span>
      </InspectorRow>
      <InspectorRow label="Bullets">
        <span className="slides-inline-tools">
          <ToolButton
            icon={List}
            label={bulleted ? 'Remove bullets' : 'Add bullets'}
            pressed={bulleted}
            disabled={element.paragraphs.length === 0}
            onClick={() => change(current => ({ ...current, paragraphs: current.paragraphs.map(paragraph => ({ ...paragraph, bullet: !bulleted })) }))}
          />
          <ToolButton icon={IndentDecrease} label="Decrease indent" disabled={element.paragraphs.length === 0} onClick={() => indent(-1)} />
          <ToolButton icon={IndentIncrease} label="Increase indent" disabled={element.paragraphs.length === 0} onClick={() => indent(1)} />
        </span>
      </InspectorRow>
    </InspectorSection>
  )
}

function Arrange({ documentId, element }: { readonly documentId: string; readonly element: SlideElement }) {
  return (
    <InspectorSection title="Arrange">
      <div className="slides-actions">
        <button type="button" className="button is-quiet" onClick={() => arrangeSelection(documentId, 'front')}><BringToFront aria-hidden="true" size={14} />Bring to front</button>
        <button type="button" className="button is-quiet" onClick={() => arrangeSelection(documentId, 'back')}><SendToBack aria-hidden="true" size={14} />Send to back</button>
        <button type="button" className="button is-quiet" onClick={() => duplicateSelection(documentId)}><Copy aria-hidden="true" size={14} />Duplicate<kbd>⌘D</kbd></button>
        <button type="button" className="button is-quiet" onClick={() => removeElement(documentId, element.id)}><Trash2 aria-hidden="true" size={14} />Delete<kbd>⌫</kbd></button>
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

export function SlidesInspector({ documentId, document, openDocument }: {
  readonly documentId: string
  readonly document: ReadySlides
  readonly openDocument: OpenDocument
}) {
  const element = selectedElement(document)
  const slide = currentSlide(document)
  const index = document.present.slides.findIndex(candidate => candidate.id === slide?.id)
  const { width, height } = document.present
  return (
    <Inspector label="Presentation inspector">
      {element !== undefined && (
        <>
          <InspectorSection title={describe(element)}>
            <p className="inspector-note">Drag to move, drag a handle to resize. Arrow keys nudge by 1 pt, with Shift by 10 pt.</p>
          </InspectorSection>
          <Placement documentId={documentId} element={element} />
          {element.kind !== 'image' && <FillAndBorder documentId={documentId} element={element} />}
          {holdsText(element) && <TextOptions documentId={documentId} element={element} />}
          <Arrange documentId={documentId} element={element} />
        </>
      )}
      {element === undefined && slide !== undefined && (
        <InspectorSection title={`Slide ${index + 1}`}>
          <InspectorRow label="Background">
            <Swatches label="Background" options={backgrounds} current={slide.background} onChoose={value => { if (value !== undefined) setSlideBackground(documentId, value) }} />
          </InspectorRow>
          <InspectorRow label="Objects"><span>{slide.elements.length}</span></InspectorRow>
        </InspectorSection>
      )}
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
