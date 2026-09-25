import { useEffect, useRef, useState, type CSSProperties, type PointerEvent as ReactPointerEvent } from 'react'
import { toast } from 'sonner'
import { encodableText } from './encodable-text'
import type { PdfPoint } from './engine'
import { markupBounds, resizeMarkup, translateMarkup, type Markup, type PageItem, type PlacedMarkup, type Size } from './model'
import { addMarkup, removeMarkup, updateMarkup } from './markup-actions'
import { dataUrlBytes, pngUrl, useSignatureStore } from './signatures'
import { select, setTool, useToolStore, type TextDraft, type Tool } from './tool-store'

interface MarkupLayerProps {
  readonly documentId: string
  readonly item: PageItem
  /** Displayed page size in points. */
  readonly size: Size
  readonly zoom: number
}

type Draft =
  | { readonly kind: 'box'; readonly start: PdfPoint; readonly end: PdfPoint }
  | { readonly kind: 'ink'; readonly points: readonly PdfPoint[] }

interface Drag {
  readonly markupId: string
  readonly mode: 'move' | 'resize'
  readonly start: PdfPoint
  readonly delta: PdfPoint
}

type TextMarkup = Extract<Markup, { kind: 'text' }>

const boxTools: ReadonlySet<Tool> = new Set<Tool>(['highlight', 'rect', 'whiteout', 'redact'])
/** Line height the engine uses for text, and how far above the CSS line box its first baseline sits. */
const lineHeight = 1.2
const baselineShift = 0.13

function boxBetween(a: PdfPoint, b: PdfPoint) {
  return { x: Math.min(a.x, b.x), y: Math.min(a.y, b.y), width: Math.abs(a.x - b.x), height: Math.abs(a.y - b.y) }
}

function boxMarkup(tool: Tool, box: ReturnType<typeof boxBetween>): Markup | undefined {
  const { inkColor, highlightColor } = useToolStore.getState()
  switch (tool) {
    case 'highlight':
      return { kind: 'rect', ...box, fill: highlightColor, opacity: 0.35 }
    case 'rect':
      return { kind: 'rect', ...box, stroke: inkColor, opacity: 1 }
    case 'whiteout':
      return { kind: 'rect', ...box, fill: '#ffffff', opacity: 1 }
    case 'redact':
      return { kind: 'redact', ...box }
    default:
      return undefined
  }
}

function applyDrag(markup: Markup, drag: Drag | undefined): Markup {
  if (drag === undefined) return markup
  if (drag.mode === 'move') return translateMarkup(markup, drag.delta.x, drag.delta.y)
  const bounds = markupBounds(markup)
  return resizeMarkup(markup, bounds.width + drag.delta.x, bounds.height + drag.delta.y)
}

function textStyle(markup: TextMarkup, zoom: number): CSSProperties {
  return {
    left: markup.x * zoom,
    top: (markup.y - markup.size * baselineShift) * zoom,
    fontSize: markup.size * zoom,
    lineHeight,
    color: markup.color,
  }
}

function blurActiveEditor(): void {
  const active = window.document.activeElement
  if (active instanceof HTMLElement) active.blur()
}

/** The markups of one page, and the pointer handling that creates, selects, moves and resizes them. */
export function MarkupLayer({ documentId, item, size, zoom }: MarkupLayerProps) {
  const overlayRef = useRef<HTMLDivElement>(null)
  const tool = useToolStore(state => state.tool)
  const selection = useToolStore(state => state.selection)
  const textDraft = useToolStore(state => state.textDraft)
  const [draft, setDraft] = useState<Draft>()
  const [drag, setDrag] = useState<Drag>()

  const pointOf = (event: ReactPointerEvent): PdfPoint => {
    const rect = overlayRef.current?.getBoundingClientRect()
    if (rect === undefined) return { x: 0, y: 0 }
    return {
      x: Math.min(size.width, Math.max(0, (event.clientX - rect.left) / zoom)),
      y: Math.min(size.height, Math.max(0, (event.clientY - rect.top) / zoom)),
    }
  }

  const placeSignature = (point: PdfPoint): void => {
    const { armedSignatureId } = useToolStore.getState()
    const signature = useSignatureStore.getState().signatures.find(candidate => candidate.id === armedSignatureId)
    if (signature === undefined) {
      toast.info('Choose a signature first', { description: 'Use the signature button in the toolbar to pick or draw one.' })
      return
    }
    const width = Math.min(160, size.width * 0.4)
    const height = width * (signature.height / signature.width)
    const id = addMarkup(documentId, item.key, {
      kind: 'image',
      x: Math.max(0, Math.min(size.width - width, point.x - width / 2)),
      y: Math.max(0, Math.min(size.height - height, point.y - height / 2)),
      width,
      height,
      png: dataUrlBytes(signature.dataUrl),
    })
    setTool('select')
    select({ pageKey: item.key, markupId: id })
  }

  const handleOverlayDown = (event: ReactPointerEvent<HTMLDivElement>): void => {
    if (event.button !== 0) return
    // The page clears the selection on pointer down; creating a markup here must not reach it.
    event.stopPropagation()
    const point = pointOf(event)
    if (tool === 'text') {
      event.preventDefault()
      if (useToolStore.getState().textDraft !== undefined) {
        blurActiveEditor()
        return
      }
      useToolStore.setState({ textDraft: { pageKey: item.key, point } })
      return
    }
    if (tool === 'signature') {
      placeSignature(point)
      return
    }
    if (boxTools.has(tool) || tool === 'ink') {
      event.currentTarget.setPointerCapture(event.pointerId)
      select(undefined)
      setDraft(tool === 'ink' ? { kind: 'ink', points: [point] } : { kind: 'box', start: point, end: point })
    }
  }

  const handleOverlayMove = (event: ReactPointerEvent<HTMLDivElement>): void => {
    if (draft === undefined) return
    const point = pointOf(event)
    setDraft(draft.kind === 'ink' ? { kind: 'ink', points: [...draft.points, point] } : { ...draft, end: point })
  }

  const handleOverlayUp = (): void => {
    if (draft === undefined) return
    setDraft(undefined)
    if (draft.kind === 'ink') {
      if (draft.points.length < 2) return
      const { inkColor } = useToolStore.getState()
      addMarkup(documentId, item.key, { kind: 'ink', points: draft.points, width: 2, color: inkColor })
      select(undefined)
      return
    }
    const box = boxBetween(draft.start, draft.end)
    if (box.width < 3 || box.height < 3) return
    const markup = boxMarkup(tool, box)
    if (markup === undefined) return
    addMarkup(documentId, item.key, markup)
    select(undefined)
  }

  const startDrag = (event: ReactPointerEvent<Element>, placed: PlacedMarkup, mode: Drag['mode']): void => {
    if (tool !== 'select' || event.button !== 0) return
    event.stopPropagation()
    event.currentTarget.setPointerCapture(event.pointerId)
    select({ pageKey: item.key, markupId: placed.id })
    setDrag({ markupId: placed.id, mode, start: pointOf(event), delta: { x: 0, y: 0 } })
  }

  const moveDrag = (event: ReactPointerEvent<Element>): void => {
    if (drag === undefined) return
    const point = pointOf(event)
    setDrag({ ...drag, delta: { x: point.x - drag.start.x, y: point.y - drag.start.y } })
  }

  const endDrag = (placed: PlacedMarkup): void => {
    if (drag === undefined) return
    const moved = applyDrag(placed.markup, drag)
    setDrag(undefined)
    if (drag.delta.x !== 0 || drag.delta.y !== 0) updateMarkup(documentId, item.key, placed.id, moved)
  }

  const dragHandlers = (placed: PlacedMarkup) => ({
    onPointerDown: (event: ReactPointerEvent<Element>) => startDrag(event, placed, 'move'),
    onPointerMove: moveDrag,
    onPointerUp: () => endDrag(placed),
    onDoubleClick: () => {
      if (placed.markup.kind === 'text') {
        useToolStore.setState({ textDraft: { pageKey: item.key, markupId: placed.id, point: placed.markup } })
      }
    },
  })

  const resizeHandle = (placed: PlacedMarkup) => (
    <span
      className="pdf-handle"
      onPointerDown={event => startDrag(event, placed, 'resize')}
      onPointerMove={moveDrag}
      onPointerUp={() => endDrag(placed)}
    />
  )

  const editingId = textDraft?.pageKey === item.key ? textDraft.markupId : undefined

  return (
    <div
      ref={overlayRef}
      className="pdf-overlay"
      data-tool={tool}
      onPointerDown={handleOverlayDown}
      onPointerMove={handleOverlayMove}
      onPointerUp={handleOverlayUp}
    >
      {item.markups.map(placed => {
        if (placed.id === editingId) return null
        const markup = applyDrag(placed.markup, drag?.markupId === placed.id ? drag : undefined)
        const selected = selection?.pageKey === item.key && selection.markupId === placed.id
        const className = `pdf-markup${selected ? ' is-selected' : ''}`
        switch (markup.kind) {
          case 'text':
            return <div key={placed.id} className={`${className} pdf-text`} style={textStyle(markup, zoom)} {...dragHandlers(placed)}>{markup.text}</div>
          case 'image':
            return (
              <div key={placed.id} className={className} style={{ left: markup.x * zoom, top: markup.y * zoom, width: markup.width * zoom, height: markup.height * zoom }} {...dragHandlers(placed)}>
                <img src={pngUrl(markup.png)} alt="" draggable={false} />
                {selected && resizeHandle(placed)}
              </div>
            )
          case 'rect':
            return (
              <div
                key={placed.id}
                className={`${className} pdf-rect${markup.fill !== undefined && markup.opacity < 1 ? ' is-highlight' : ''}`}
                style={{
                  left: markup.x * zoom,
                  top: markup.y * zoom,
                  width: markup.width * zoom,
                  height: markup.height * zoom,
                  background: markup.fill,
                  opacity: markup.opacity,
                  borderColor: markup.stroke,
                  borderWidth: markup.stroke === undefined ? 0 : Math.max(1, zoom),
                }}
                {...dragHandlers(placed)}
              >
                {selected && resizeHandle(placed)}
              </div>
            )
          case 'redact':
            return (
              <div key={placed.id} className={`${className} pdf-redact`} style={{ left: markup.x * zoom, top: markup.y * zoom, width: markup.width * zoom, height: markup.height * zoom }} {...dragHandlers(placed)}>
                {selected && resizeHandle(placed)}
              </div>
            )
          case 'ink': {
            const bounds = markupBounds(markup)
            return (
              <div key={placed.id} className="pdf-ink-wrap">
                <svg className="pdf-ink" viewBox={`0 0 ${size.width} ${size.height}`} style={{ width: size.width * zoom, height: size.height * zoom }}>
                  <polyline className="pdf-ink-hit" points={markup.points.map(point => `${point.x},${point.y}`).join(' ')} {...dragHandlers(placed)} />
                  <polyline points={markup.points.map(point => `${point.x},${point.y}`).join(' ')} stroke={markup.color} strokeWidth={markup.width} />
                </svg>
                {selected && <div className="pdf-markup is-selected pdf-bounds" style={{ left: bounds.x * zoom, top: bounds.y * zoom, width: bounds.width * zoom, height: bounds.height * zoom }} />}
              </div>
            )
          }
        }
      })}

      {draft?.kind === 'box' && (() => {
        const box = boxBetween(draft.start, draft.end)
        return <div className={`pdf-draft pdf-draft--${tool}`} style={{ left: box.x * zoom, top: box.y * zoom, width: box.width * zoom, height: box.height * zoom }} />
      })()}
      {draft?.kind === 'ink' && (
        <svg className="pdf-ink" viewBox={`0 0 ${size.width} ${size.height}`} style={{ width: size.width * zoom, height: size.height * zoom }}>
          <polyline points={draft.points.map(point => `${point.x},${point.y}`).join(' ')} stroke={useToolStore.getState().inkColor} strokeWidth={2} />
        </svg>
      )}

      {textDraft?.pageKey === item.key && (
        <TextEditor
          key={textDraft.markupId ?? `${textDraft.point.x},${textDraft.point.y}`}
          documentId={documentId}
          draft={textDraft}
          existing={item.markups.find(placed => placed.id === textDraft.markupId)?.markup}
          zoom={zoom}
        />
      )}
    </div>
  )
}

interface TextEditorProps {
  readonly documentId: string
  readonly draft: TextDraft
  readonly existing: Markup | undefined
  readonly zoom: number
}

/**
 * Inline text entry. Commits on blur or ⌘Return, and on unmount (switching tabs from the keyboard
 * removes the editor without a blur); Escape discards. An emptied existing box is removed.
 */
function TextEditor({ documentId, draft, existing, zoom }: TextEditorProps) {
  const current = existing?.kind === 'text' ? existing : undefined
  const { textSize, inkColor } = useToolStore.getState()
  const [value, setValue] = useState(current?.text ?? '')
  const typed = useRef(value)
  const finished = useRef(false)
  const style = textStyle(current ?? { kind: 'text', x: draft.point.x, y: draft.point.y, text: '', size: textSize, color: inkColor }, zoom)

  /** Writes the typed text to the page once; returns the id of a new markup. */
  const commit = (): string | undefined => {
    if (finished.current) return undefined
    finished.current = true
    const text = encodableText(typed.current.replace(/\t/g, '  ')).replace(/\s+$/, '')
    if (current !== undefined && draft.markupId !== undefined) {
      if (text === '') removeMarkup(documentId, draft.pageKey, draft.markupId)
      else if (text !== current.text) updateMarkup(documentId, draft.pageKey, draft.markupId, { ...current, text })
      return undefined
    }
    if (text === '') return undefined
    return addMarkup(documentId, draft.pageKey, { kind: 'text', x: draft.point.x, y: draft.point.y, text, size: textSize, color: inkColor })
  }

  const finish = (keep: boolean): void => {
    if (!keep) finished.current = true
    const id = commit()
    useToolStore.setState({ textDraft: undefined })
    if (id !== undefined) select({ pageKey: draft.pageKey, markupId: id })
  }

  useEffect(() => {
    // Re-armed on mount because development mode mounts twice; the first unmount has nothing typed.
    finished.current = false
    return () => void commit()
  }, [])

  return (
    <textarea
      autoFocus
      className="pdf-text pdf-text-editor"
      style={style}
      value={value}
      rows={1}
      spellCheck
      aria-label="Text"
      onChange={event => {
        typed.current = event.currentTarget.value
        setValue(event.currentTarget.value)
      }}
      onBlur={() => finish(true)}
      onPointerDown={event => event.stopPropagation()}
      onKeyDown={event => {
        event.stopPropagation()
        if (event.key === 'Escape') {
          event.preventDefault()
          finish(false)
        } else if (event.key === 'Enter' && event.metaKey) {
          event.preventDefault()
          finish(true)
        }
      }}
    />
  )
}
