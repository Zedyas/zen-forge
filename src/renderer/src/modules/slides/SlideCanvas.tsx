import { useEffect, useLayoutEffect, useState, type CSSProperties, type PointerEvent as ReactPointerEvent } from 'react'
import { ContextMenu } from '@base-ui/react/context-menu'
import {
  boundsOf,
  canRotate,
  findElement,
  fitTable,
  holdsText,
  isLine,
  lineEnds,
  lineThrough,
  plainText,
  resizeFrame,
  translateElement,
  type Box,
  type Handle,
  type Point,
  type SlideElement,
  type TableElement,
} from './model'
import { setCellText } from './format-actions'
import {
  arrangeSelection,
  duplicateSelection,
  removeSelection,
  replaceElements,
  select,
  setCurrentCell,
  startEditing,
  toggleSelected,
} from './slides-actions'
import { currentSlide, selectedElements, updateSlides, type ReadySlides, type TableCellAddress } from './slides-store'
import { fontStack, SlideView } from './SlideView'
import { TextEditor } from './TextEditor'

type Drag =
  | { readonly kind: 'move'; readonly origins: readonly SlideElement[]; readonly start: Point; readonly current: readonly SlideElement[]; readonly guides: Guides }
  | { readonly kind: 'resize'; readonly origin: SlideElement; readonly handle: Handle; readonly start: Point; readonly current: SlideElement }
  | { readonly kind: 'end'; readonly origin: SlideElement; readonly end: 0 | 1; readonly start: Point; readonly current: SlideElement }
  | { readonly kind: 'rotate'; readonly origin: SlideElement; readonly centre: Point; readonly current: SlideElement }
  /** Selecting by dragging a box over empty slide; points are on the slide. */
  | { readonly kind: 'marquee'; readonly start: Point; readonly end: Point; readonly additive: boolean }

/** Slide lines a moved selection snapped to, in points: x positions of vertical lines, y of horizontal ones. */
interface Guides {
  readonly x: readonly number[]
  readonly y: readonly number[]
}

const noGuides: Guides = { x: [], y: [] }

const handles: readonly Handle[] = [
  { h: -1, v: -1 }, { h: 0, v: -1 }, { h: 1, v: -1 },
  { h: -1, v: 0 }, { h: 1, v: 0 },
  { h: -1, v: 1 }, { h: 0, v: 1 }, { h: 1, v: 1 },
]

const handleCursors: Record<string, string> = {
  '-1,-1': 'nwse-resize', '1,1': 'nwse-resize', '1,-1': 'nesw-resize', '-1,1': 'nesw-resize',
  '0,-1': 'ns-resize', '0,1': 'ns-resize', '-1,0': 'ew-resize', '1,0': 'ew-resize',
}

/** How close, in screen pixels, an edge or centre must come to the slide's to snap to it. */
const snapDistance = 6

/**
 * Moves elements by `delta`, snapping the selection's edges and centre to the slide's edges and
 * centre lines when they come close, as Keynote's guides do.
 */
function snappedMove(origins: readonly SlideElement[], delta: Point, slide: Box, threshold: number): { readonly moved: SlideElement[]; readonly guides: Guides } {
  const bounds = boundsOf(origins)
  if (bounds === undefined) return { moved: [], guides: noGuides }
  const snap = (start: number, size: number, extent: number): { readonly offset: number; readonly line?: number } => {
    let best: { offset: number; line?: number } = { offset: 0 }
    let distance = threshold
    for (const line of [0, extent / 2, extent]) {
      for (const edge of [start, start + size / 2, start + size]) {
        if (Math.abs(line - edge) < distance) {
          distance = Math.abs(line - edge)
          best = { offset: line - edge, line }
        }
      }
    }
    return best
  }
  const x = snap(bounds.x + delta.x, bounds.width, slide.width)
  const y = snap(bounds.y + delta.y, bounds.height, slide.height)
  return {
    moved: origins.map(origin => translateElement(origin, delta.x + x.offset, delta.y + y.offset)),
    guides: { x: x.line === undefined ? [] : [x.line], y: y.line === undefined ? [] : [y.line] },
  }
}

/** A single element's drag: resizing (Shift keeps proportions; pictures always keep them from a corner), a line's end, or turning. */
function draggedOne(drag: Exclude<Drag, { kind: 'move' | 'marquee' }>, delta: Point, pointer: Point, shift: boolean): SlideElement {
  const { origin } = drag
  switch (drag.kind) {
    case 'resize': {
      const corner = drag.handle.h !== 0 && drag.handle.v !== 0
      const resized = resizeFrame(origin, drag.handle, delta, corner && (shift || origin.kind === 'image'))
      return resized.kind === 'table' ? { ...fitTable(resized, resized.width, resized.height), x: resized.x, y: resized.y } : resized
    }
    case 'end': {
      if (origin.kind !== 'shape') return origin
      const ends = [...lineEnds(origin)]
      const moved = ends[drag.end]
      if (moved === undefined) return origin
      ends[drag.end] = { x: moved.x + delta.x, y: moved.y + delta.y }
      const [start, end] = ends
      return start === undefined || end === undefined ? origin : lineThrough(origin, start, end)
    }
    case 'rotate': {
      const angle = (Math.atan2(pointer.y - drag.centre.y, pointer.x - drag.centre.x) * 180) / Math.PI + 90
      const step = shift ? 15 : 1
      return { ...origin, rotation: ((Math.round(angle / step) * step) % 360 + 360) % 360 }
    }
  }
}

function screenFrame(element: Box & { readonly rotation?: number }, scale: number): CSSProperties {
  return {
    left: element.x * scale,
    top: element.y * scale,
    width: element.width * scale,
    height: element.height * scale,
    transform: element.rotation === undefined || element.rotation === 0 ? undefined : `rotate(${element.rotation}deg)`,
  }
}

/** Fits the slide into the space left by the rail, inspector and notes. */
function useFittedScale(desk: HTMLElement | null, width: number, height: number): number {
  const [scale, setScale] = useState(0)
  useEffect(() => {
    if (desk === null) return
    const fit = (): void => {
      const pad = 40
      setScale(Math.max(0.05, Math.min((desk.clientWidth - pad * 2) / width, (desk.clientHeight - pad * 2) / height)))
    }
    fit()
    const observer = new ResizeObserver(fit)
    observer.observe(desk)
    return () => observer.disconnect()
  }, [desk, width, height])
  return scale
}

/** The table cell under a point on screen, read from the drawn table. */
function cellAt(tableId: string, clientX: number, clientY: number): TableCellAddress | undefined {
  const cell = window.document.elementsFromPoint(clientX, clientY).find(element =>
    element instanceof HTMLElement && element.matches('td[data-row]') && element.closest(`[data-element-id="${tableId}"]`) !== null)
  if (!(cell instanceof HTMLElement)) return undefined
  return { row: Number(cell.dataset['row']), column: Number(cell.dataset['column']) }
}

function sameBox(a: Box | undefined, b: Box | undefined): boolean {
  return a === b || (a !== undefined && b !== undefined && a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height)
}

/** Where a table cell is drawn, relative to the stage, in screen pixels; measured after each drawing of the slide. */
function useCellRect(stage: HTMLElement | null, tableId: string | undefined, cell: TableCellAddress | undefined, drawn: unknown, scale: number): Box | undefined {
  const [rect, setRect] = useState<Box>()
  const row = cell?.row
  const column = cell?.column
  useLayoutEffect(() => {
    const element = tableId === undefined || row === undefined || stage === null ? null
      : stage.querySelector(`.slide-view [data-element-id="${tableId}"] td[data-row="${row}"][data-column="${column}"]`)
    let next: Box | undefined
    if (element !== null && stage !== null) {
      const outer = stage.getBoundingClientRect()
      const inner = element.getBoundingClientRect()
      next = { x: inner.left - outer.left, y: inner.top - outer.top, width: inner.width, height: inner.height }
    }
    setRect(previous => sameBox(previous, next) ? previous : next)
  }, [stage, tableId, row, column, drawn, scale])
  return rect
}

interface CellEditorProps {
  readonly documentId: string
  readonly table: TableElement
  readonly cell: TableCellAddress
  readonly rect: Box
  readonly scale: number
}

/**
 * Types one table cell's text; Tab and Shift-Tab move to the next and previous cell, Escape closes.
 * The field grows with its text, and so does the row, as in PowerPoint.
 */
function CellEditor({ documentId, table, cell, rect, scale }: CellEditorProps) {
  const source = table.rows[cell.row]?.cells[cell.column]
  const [text, setText] = useState(source?.text ?? '')
  const [height, setHeight] = useState(rect.height)
  // The field is drawn at the slide's scale, so its height in points is its pixel height over the scale.
  // It only asks for a taller row once the text outgrew the cell.
  const write = (): void => setCellText(documentId, table.id, cell, text, height > rect.height + 1 ? height / scale : 0)
  const go = (step: 1 | -1): void => {
    write()
    const columns = table.columns.length
    const index = Math.min(table.rows.length * columns - 1, Math.max(0, cell.row * columns + cell.column + step))
    startEditing(documentId, table.id, { row: Math.floor(index / columns), column: index % columns })
  }
  return (
    <textarea
      autoFocus
      className="slides-cell-editor"
      value={text}
      aria-label={`Row ${cell.row + 1}, column ${cell.column + 1}`}
      style={{
        ...screenFrame({ ...rect, height }, 1),
        padding: `${3.6 * scale}px ${7.2 * scale}px`,
        fontFamily: fontStack(table.font),
        fontSize: table.size * scale,
        fontWeight: source?.bold === true ? 700 : 400,
        color: source?.color ?? table.color,
        background: source?.fill ?? '#ffffff',
        textAlign: source?.align,
      }}
      onChange={event => {
        const field = event.currentTarget
        setText(field.value)
        field.style.height = '0px'
        setHeight(Math.max(rect.height, field.scrollHeight))
        field.style.height = ''
      }}
      onPointerDown={event => event.stopPropagation()}
      onBlur={() => {
        write()
        // Moving to another cell has already opened that one; only a click elsewhere closes the editor.
        updateSlides(documentId, current => current.editingId === table.id && current.cell?.row === cell.row && current.cell.column === cell.column ? { editingId: undefined } : {})
      }}
      onKeyDown={event => {
        event.stopPropagation()
        if (event.key === 'Tab') {
          event.preventDefault()
          go(event.shiftKey ? -1 : 1)
        } else if (event.key === 'Escape') {
          event.preventDefault()
          event.currentTarget.blur()
        }
      }}
    />
  )
}

interface SlideCanvasProps {
  readonly documentId: string
  readonly document: ReadySlides
}

/** The current slide, fitted to the space, with selection, moving, resizing, turning and inline text editing. */
export function SlideCanvas({ documentId, document }: SlideCanvasProps) {
  const [desk, setDesk] = useState<HTMLElement | null>(null)
  const [stage, setStage] = useState<HTMLElement | null>(null)
  const [drag, setDrag] = useState<Drag>()
  const [caret, setCaret] = useState<Point>()
  const { width, height } = document.present
  const scale = useFittedScale(desk, width, height)
  const slide = currentSlide(document)
  const selected = selectedElements(document)
  const single = selected.length === 1 ? selected[0] : undefined
  const editing = findElement(slide, document.editingId)
  const table = single?.kind === 'table' ? single : undefined
  const cellRect = useCellRect(stage, table?.id, document.cell, slide, scale)

  if (slide === undefined) return <div ref={setDesk} className="slides-desk" />

  const changed = drag === undefined ? [] : drag.kind === 'move' ? drag.current : drag.kind === 'marquee' ? [] : [drag.current]
  const byId = new Map(changed.map(element => [element.id, element]))
  const shown = byId.size === 0 ? slide : { ...slide, elements: slide.elements.map(element => byId.get(element.id) ?? element) }
  const shownSelected = selected.map(element => byId.get(element.id) ?? element)

  const slidePoint = (event: { readonly clientX: number; readonly clientY: number }): Point => {
    const rect = stage?.getBoundingClientRect()
    return rect === undefined ? { x: 0, y: 0 } : { x: (event.clientX - rect.left) / scale, y: (event.clientY - rect.top) / scale }
  }

  const beginMove = (event: ReactPointerEvent, element: SlideElement): void => {
    if (event.button === 2) {
      if (!document.selection.includes(element.id)) select(documentId, [element.id])
      return
    }
    if (event.button !== 0) return
    event.stopPropagation()
    if (event.shiftKey) {
      toggleSelected(documentId, element.id)
      return
    }
    const already = document.selection.includes(element.id)
    const ids = already ? document.selection : [element.id]
    if (!already) select(documentId, [element.id])
    if (element.kind === 'table' && already) setCurrentCell(documentId, cellAt(element.id, event.clientX, event.clientY))
    event.currentTarget.setPointerCapture(event.pointerId)
    const origins = slide.elements.filter(candidate => ids.includes(candidate.id))
    setDrag({ kind: 'move', origins, start: { x: event.clientX, y: event.clientY }, current: origins, guides: noGuides })
  }

  const beginOne = (event: ReactPointerEvent, next: Exclude<Drag, { kind: 'move' | 'marquee' }>): void => {
    if (event.button !== 0) return
    event.stopPropagation()
    event.currentTarget.setPointerCapture(event.pointerId)
    setDrag(next)
  }

  const continueDrag = (event: ReactPointerEvent): void => {
    if (drag === undefined) return
    if (drag.kind === 'marquee') {
      setDrag({ ...drag, end: slidePoint(event) })
      return
    }
    if (drag.kind === 'rotate') {
      setDrag({ ...drag, current: draggedOne(drag, { x: 0, y: 0 }, { x: event.clientX, y: event.clientY }, event.shiftKey) })
      return
    }
    const delta = { x: (event.clientX - drag.start.x) / scale, y: (event.clientY - drag.start.y) / scale }
    if (drag.kind === 'move') {
      // Shift moves along one axis only.
      const straight = event.shiftKey ? (Math.abs(delta.x) > Math.abs(delta.y) ? { x: delta.x, y: 0 } : { x: 0, y: delta.y }) : delta
      const { moved, guides } = snappedMove(drag.origins, straight, { x: 0, y: 0, width, height }, snapDistance / scale)
      setDrag({ ...drag, current: moved, guides })
      return
    }
    setDrag({ ...drag, current: draggedOne(drag, delta, { x: event.clientX, y: event.clientY }, event.shiftKey) })
  }

  const endDrag = (): void => {
    if (drag === undefined) return
    setDrag(undefined)
    if (drag.kind === 'marquee') {
      const box = { x: Math.min(drag.start.x, drag.end.x), y: Math.min(drag.start.y, drag.end.y), width: Math.abs(drag.end.x - drag.start.x), height: Math.abs(drag.end.y - drag.start.y) }
      if (box.width * scale < 3 && box.height * scale < 3) return
      const inside = slide.elements.filter(element =>
        element.x < box.x + box.width && element.x + element.width > box.x && element.y < box.y + box.height && element.y + element.height > box.y)
      select(documentId, [...(drag.additive ? document.selection : []), ...inside.map(element => element.id).filter(id => !drag.additive || !document.selection.includes(id))])
      return
    }
    const origins = drag.kind === 'move' ? drag.origins : [drag.origin]
    const current = drag.kind === 'move' ? drag.current : [drag.current]
    const moved = current.some((element, index) => JSON.stringify(element) !== JSON.stringify(origins[index]))
    if (moved) replaceElements(documentId, current)
  }

  // Only the desk listens: a pointer captured by a hit area or handle still bubbles its moves up here.
  const dragEvents = { onPointerMove: continueDrag, onPointerUp: endDrag, onPointerCancel: endDrag }

  const stageRect = stage?.getBoundingClientRect()
  const centreOnScreen = (element: SlideElement): Point => ({
    x: (stageRect?.left ?? 0) + (element.x + element.width / 2) * scale,
    y: (stageRect?.top ?? 0) + (element.y + element.height / 2) * scale,
  })

  const guides = drag?.kind === 'move' ? drag.guides : noGuides
  const marquee = drag?.kind === 'marquee'
    ? { x: Math.min(drag.start.x, drag.end.x), y: Math.min(drag.start.y, drag.end.y), width: Math.abs(drag.end.x - drag.start.x), height: Math.abs(drag.end.y - drag.start.y) }
    : undefined
  const shownSingle = shownSelected.length === 1 ? shownSelected[0] : undefined

  return (
    <div
      ref={setDesk}
      className="slides-desk"
      onPointerDown={event => {
        if (event.button !== 0 || !(event.target instanceof Element)) return
        if (event.target.closest('.slides-hit, .slides-handle, .slide-text-editor, .slides-cell-editor, .slides-editor-layer .slide-text-frame') !== null) return
        if (!event.shiftKey) select(documentId, [])
        event.currentTarget.setPointerCapture(event.pointerId)
        const point = slidePoint(event)
        setDrag({ kind: 'marquee', start: point, end: point, additive: event.shiftKey })
      }}
      {...dragEvents}
    >
      {scale > 0 && (
        <ContextMenu.Root>
          <ContextMenu.Trigger ref={setStage} className="slides-stage" style={{ width: width * scale, height: height * scale }}>
            <SlideView slide={shown} width={width} height={height} scale={scale} hiddenId={editing !== undefined && holdsText(editing) ? editing.id : undefined} />

            <div className="slides-hits">
              {shown.elements.map(element => {
                if (element.id === editing?.id && holdsText(element)) return null
                if (isLine(element)) {
                  const [start, end] = lineEnds(element)
                  return (
                    <svg key={element.id} className="slides-line-hit" width={width * scale} height={height * scale}>
                      <line className="slides-hit" x1={start.x * scale} y1={start.y * scale} x2={end.x * scale} y2={end.y * scale} onPointerDown={event => beginMove(event, element)} />
                    </svg>
                  )
                }
                const empty = element.kind === 'text' && plainText(element.paragraphs).trim() === ''
                return (
                  <div
                    key={element.id}
                    className={`slides-hit${empty ? ' is-empty' : ''}`}
                    style={screenFrame(element, scale)}
                    onPointerDown={event => beginMove(event, element)}
                    onDoubleClick={event => {
                      if (element.kind === 'table') {
                        startEditing(documentId, element.id, cellAt(element.id, event.clientX, event.clientY) ?? { row: 0, column: 0 })
                      } else if (holdsText(element)) {
                        setCaret({ x: event.clientX, y: event.clientY })
                        startEditing(documentId, element.id)
                      }
                    }}
                  >
                    {empty && element.kind === 'text' && element.prompt !== undefined && (
                      <span className="slides-prompt" style={{ fontSize: (element.paragraphs[0]?.runs[0]?.size ?? 18) * scale * 0.8, textAlign: element.paragraphs[0]?.align }}>
                        {element.prompt}
                      </span>
                    )}
                  </div>
                )
              })}
            </div>

            {guides.x.map(x => <span key={`x${x}`} className="slides-guide is-vertical" style={{ left: x * scale }} />)}
            {guides.y.map(y => <span key={`y${y}`} className="slides-guide is-horizontal" style={{ top: y * scale }} />)}
            {marquee !== undefined && <span className="slides-marquee" style={screenFrame(marquee, scale)} />}

            {editing === undefined && shownSelected.length > 1 && shownSelected.map(element => (
              <div key={element.id} className="slides-selection" style={screenFrame(element, scale)} />
            ))}

            {editing === undefined && shownSingle !== undefined && single !== undefined && (
              isLine(shownSingle) ? (
                lineEnds(shownSingle).map((point, end) => (
                  <span
                    key={end}
                    className="slides-handle is-end"
                    style={{ left: point.x * scale, top: point.y * scale }}
                    onPointerDown={event => beginOne(event, { kind: 'end', origin: single, end: end === 0 ? 0 : 1, start: { x: event.clientX, y: event.clientY }, current: single })}
                  />
                ))
              ) : (
                <div className="slides-selection" style={screenFrame(shownSingle, scale)}>
                  {handles.map(handle => (
                    <span
                      key={`${handle.h},${handle.v}`}
                      className="slides-handle"
                      style={{ left: `${(handle.h + 1) * 50}%`, top: `${(handle.v + 1) * 50}%`, cursor: handleCursors[`${handle.h},${handle.v}`] }}
                      onPointerDown={event => beginOne(event, { kind: 'resize', origin: single, handle, start: { x: event.clientX, y: event.clientY }, current: single })}
                    />
                  ))}
                  {canRotate(shownSingle) && (
                    <span
                      className="slides-handle is-rotate"
                      aria-label="Rotate"
                      onPointerDown={event => beginOne(event, { kind: 'rotate', origin: single, centre: centreOnScreen(single), current: single })}
                    />
                  )}
                </div>
              )
            )}

            {table !== undefined && cellRect !== undefined && editing === undefined && document.cell !== undefined && (
              <span className="slides-cell" style={screenFrame(cellRect, 1)} />
            )}

            {editing !== undefined && holdsText(editing) && (
              <div className="slides-editor-layer" style={{ width, height, transform: `scale(${scale})` }}>
                <TextEditor key={editing.id} documentId={documentId} element={editing} caret={caret} />
              </div>
            )}
            {editing?.kind === 'table' && document.cell !== undefined && cellRect !== undefined && (
              <CellEditor
                key={`${editing.id}:${document.cell.row}:${document.cell.column}`}
                documentId={documentId}
                table={editing}
                cell={document.cell}
                rect={cellRect}
                scale={scale}
              />
            )}
          </ContextMenu.Trigger>
          <ContextMenu.Portal>
            <ContextMenu.Positioner>
              <ContextMenu.Popup className="menu-popup">
                <ContextMenu.Item className="menu-item" disabled={selected.length === 0} onClick={() => duplicateSelection(documentId)}>
                  <span>Duplicate</span><kbd>⌘D</kbd>
                </ContextMenu.Item>
                <ContextMenu.Separator className="menu-separator" />
                <ContextMenu.Item className="menu-item" disabled={selected.length === 0} onClick={() => arrangeSelection(documentId, 'front')}><span>Bring to Front</span></ContextMenu.Item>
                <ContextMenu.Item className="menu-item" disabled={selected.length === 0} onClick={() => arrangeSelection(documentId, 'back')}><span>Send to Back</span></ContextMenu.Item>
                <ContextMenu.Separator className="menu-separator" />
                <ContextMenu.Item className="menu-item" disabled={selected.length === 0} onClick={() => removeSelection(documentId)}>
                  <span>Delete</span><kbd>⌫</kbd>
                </ContextMenu.Item>
              </ContextMenu.Popup>
            </ContextMenu.Positioner>
          </ContextMenu.Portal>
        </ContextMenu.Root>
      )}
    </div>
  )
}
