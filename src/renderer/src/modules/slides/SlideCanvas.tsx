import { useEffect, useState, type CSSProperties, type PointerEvent as ReactPointerEvent } from 'react'
import { ContextMenu } from '@base-ui/react/context-menu'
import {
  findElement,
  holdsText,
  isLine,
  lineEnds,
  lineThrough,
  plainText,
  resizeFrame,
  translateElement,
  updateElement,
  type Handle,
  type Point,
  type SlideElement,
} from './model'
import { arrangeSelection, changeElement, duplicateSelection, removeElement, select, startEditing } from './slides-actions'
import { currentSlide, type ReadySlides } from './slides-store'
import { SlideView } from './SlideView'
import { TextEditor } from './TextEditor'

type DragMode =
  | { readonly kind: 'move' }
  | { readonly kind: 'resize'; readonly handle: Handle }
  | { readonly kind: 'end'; readonly end: 0 | 1 }

interface Drag {
  readonly mode: DragMode
  readonly origin: SlideElement
  /** Pointer position where the drag started, in window coordinates. */
  readonly start: Point
  readonly current: SlideElement
}

const handles: readonly Handle[] = [
  { h: -1, v: -1 }, { h: 0, v: -1 }, { h: 1, v: -1 },
  { h: -1, v: 0 }, { h: 1, v: 0 },
  { h: -1, v: 1 }, { h: 0, v: 1 }, { h: 1, v: 1 },
]

const handleCursors: Record<string, string> = {
  '-1,-1': 'nwse-resize', '1,1': 'nwse-resize', '1,-1': 'nesw-resize', '-1,1': 'nesw-resize',
  '0,-1': 'ns-resize', '0,1': 'ns-resize', '-1,0': 'ew-resize', '1,0': 'ew-resize',
}

/** Where a drag has taken an element so far. Shift keeps a picture's proportions free, or a shape's fixed, as in Keynote. */
function dragged(drag: Drag, delta: Point, shift: boolean): SlideElement {
  const { origin, mode } = drag
  switch (mode.kind) {
    case 'move': {
      // Shift moves along one axis only.
      const straight = shift ? (Math.abs(delta.x) > Math.abs(delta.y) ? { x: delta.x, y: 0 } : { x: 0, y: delta.y }) : delta
      return translateElement(origin, straight.x, straight.y)
    }
    case 'resize': {
      const corner = mode.handle.h !== 0 && mode.handle.v !== 0
      const keepAspect = corner && (origin.kind === 'image' ? !shift : shift)
      return resizeFrame(origin, mode.handle, delta, keepAspect)
    }
    case 'end': {
      if (origin.kind !== 'shape') return origin
      const ends = [...lineEnds(origin)]
      const moved = ends[mode.end]
      if (moved === undefined) return origin
      ends[mode.end] = { x: moved.x + delta.x, y: moved.y + delta.y }
      const [start, end] = ends
      return start === undefined || end === undefined ? origin : lineThrough(origin, start, end)
    }
  }
}

function screenFrame(element: SlideElement, scale: number): CSSProperties {
  return {
    left: element.x * scale,
    top: element.y * scale,
    width: element.width * scale,
    height: element.height * scale,
    transform: element.rotation === 0 ? undefined : `rotate(${element.rotation}deg)`,
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

interface SlideCanvasProps {
  readonly documentId: string
  readonly document: ReadySlides
}

/** The current slide, fitted to the space, with selection, moving, resizing and inline text editing. */
export function SlideCanvas({ documentId, document }: SlideCanvasProps) {
  const [desk, setDesk] = useState<HTMLElement | null>(null)
  const [drag, setDrag] = useState<Drag>()
  const [caret, setCaret] = useState<Point>()
  const { width, height } = document.present
  const scale = useFittedScale(desk, width, height)
  const slide = currentSlide(document)
  const selected = findElement(slide, document.selectedId)
  const editing = findElement(slide, document.editingId)

  if (slide === undefined) return <div ref={setDesk} className="slides-desk" />
  const shown = drag === undefined ? slide : updateElement(slide, drag.origin.id, () => drag.current)
  const shownSelected = selected === undefined ? undefined : findElement(shown, selected.id)

  const beginDrag = (event: ReactPointerEvent, element: SlideElement, mode: DragMode): void => {
    if (event.button !== 0) {
      if (event.button === 2) select(documentId, element.id)
      return
    }
    event.stopPropagation()
    event.currentTarget.setPointerCapture(event.pointerId)
    select(documentId, element.id)
    setDrag({ mode, origin: element, start: { x: event.clientX, y: event.clientY }, current: element })
  }

  const continueDrag = (event: ReactPointerEvent): void => {
    if (drag === undefined) return
    const delta = { x: (event.clientX - drag.start.x) / scale, y: (event.clientY - drag.start.y) / scale }
    setDrag({ ...drag, current: dragged(drag, delta, event.shiftKey) })
  }

  const endDrag = (): void => {
    if (drag === undefined) return
    setDrag(undefined)
    const { origin, current } = drag
    const moved = Math.abs(current.x - origin.x) + Math.abs(current.y - origin.y) + Math.abs(current.width - origin.width) + Math.abs(current.height - origin.height) > 0.01
    if (moved) changeElement(documentId, origin.id, () => current)
  }

  const dragProps = (element: SlideElement, mode: DragMode) => ({
    onPointerDown: (event: ReactPointerEvent) => beginDrag(event, element, mode),
    onPointerMove: continueDrag,
    onPointerUp: endDrag,
    onPointerCancel: endDrag,
  })

  return (
    <div
      ref={setDesk}
      className="slides-desk"
      onPointerDown={event => {
        if (event.target instanceof Element && event.target.closest('.slides-hit, .slides-handle, .slide-text-editor') === null) select(documentId, undefined)
      }}
    >
      {scale > 0 && (
        <ContextMenu.Root>
          <ContextMenu.Trigger className="slides-stage" style={{ width: width * scale, height: height * scale }}>
            <SlideView slide={shown} width={width} height={height} scale={scale} hiddenId={editing?.id} />

            <div className="slides-hits">
              {shown.elements.map(element => {
                if (element.id === editing?.id) return null
                if (isLine(element)) {
                  const [start, end] = lineEnds(element)
                  return (
                    <svg key={element.id} className="slides-line-hit" width={width * scale} height={height * scale}>
                      <line className="slides-hit" x1={start.x * scale} y1={start.y * scale} x2={end.x * scale} y2={end.y * scale} {...dragProps(element, { kind: 'move' })} />
                    </svg>
                  )
                }
                const empty = element.kind === 'text' && plainText(element.paragraphs).trim() === ''
                return (
                  <div
                    key={element.id}
                    className={`slides-hit${empty ? ' is-empty' : ''}`}
                    style={screenFrame(element, scale)}
                    {...dragProps(element, { kind: 'move' })}
                    onDoubleClick={event => {
                      if (!holdsText(element)) return
                      setCaret({ x: event.clientX, y: event.clientY })
                      startEditing(documentId, element.id)
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

            {shownSelected !== undefined && editing === undefined && (
              isLine(shownSelected) ? (
                lineEnds(shownSelected).map((point, end) => (
                  <span
                    key={end}
                    className="slides-handle is-end"
                    style={{ left: point.x * scale, top: point.y * scale }}
                    {...dragProps(selected ?? shownSelected, { kind: 'end', end: end === 0 ? 0 : 1 })}
                  />
                ))
              ) : (
                <div className="slides-selection" style={screenFrame(shownSelected, scale)}>
                  {handles.map(handle => (
                    <span
                      key={`${handle.h},${handle.v}`}
                      className="slides-handle"
                      style={{ left: `${(handle.h + 1) * 50}%`, top: `${(handle.v + 1) * 50}%`, cursor: handleCursors[`${handle.h},${handle.v}`] }}
                      {...dragProps(selected ?? shownSelected, { kind: 'resize', handle })}
                    />
                  ))}
                </div>
              )
            )}

            {editing !== undefined && holdsText(editing) && (
              <div className="slides-editor-layer" style={{ width, height, transform: `scale(${scale})` }}>
                <TextEditor key={editing.id} documentId={documentId} element={editing} caret={caret} />
              </div>
            )}
          </ContextMenu.Trigger>
          <ContextMenu.Portal>
            <ContextMenu.Positioner>
              <ContextMenu.Popup className="menu-popup">
                <ContextMenu.Item className="menu-item" disabled={selected === undefined} onClick={() => duplicateSelection(documentId)}>
                  <span>Duplicate</span><kbd>⌘D</kbd>
                </ContextMenu.Item>
                <ContextMenu.Separator className="menu-separator" />
                <ContextMenu.Item className="menu-item" disabled={selected === undefined} onClick={() => arrangeSelection(documentId, 'front')}><span>Bring to Front</span></ContextMenu.Item>
                <ContextMenu.Item className="menu-item" disabled={selected === undefined} onClick={() => arrangeSelection(documentId, 'back')}><span>Send to Back</span></ContextMenu.Item>
                <ContextMenu.Separator className="menu-separator" />
                <ContextMenu.Item className="menu-item" disabled={selected === undefined} onClick={() => { if (selected !== undefined) removeElement(documentId, selected.id) }}>
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
