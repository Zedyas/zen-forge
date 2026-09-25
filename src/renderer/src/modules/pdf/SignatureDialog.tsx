import { Dialog } from '@base-ui/react/dialog'
import { useEffect, useRef, useState, type PointerEvent as ReactPointerEvent } from 'react'
import { newId } from './model'
import { addSignature } from './signatures'
import { setTool, useToolStore } from './tool-store'

const padWidth = 460
const padHeight = 170
const inks = [
  { label: 'Dark blue', value: '#1c2f7a' },
  { label: 'Black', value: '#1b1b1b' },
] as const

interface Stroke {
  readonly color: string
  readonly points: readonly { readonly x: number; readonly y: number }[]
}

function draw(canvas: HTMLCanvasElement, strokes: readonly Stroke[]): void {
  const context = canvas.getContext('2d')
  if (context === null) return
  const ratio = window.devicePixelRatio
  context.setTransform(ratio, 0, 0, ratio, 0, 0)
  context.clearRect(0, 0, padWidth, padHeight)
  context.lineCap = 'round'
  context.lineJoin = 'round'
  context.lineWidth = 2.4
  for (const stroke of strokes) {
    context.strokeStyle = stroke.color
    context.beginPath()
    stroke.points.forEach((point, index) => {
      if (index === 0) context.moveTo(point.x, point.y)
      else {
        // Midpoint quadratic curves smooth the pointer samples into a pen-like line.
        const previous = stroke.points[index - 1] ?? point
        context.quadraticCurveTo(previous.x, previous.y, (previous.x + point.x) / 2, (previous.y + point.y) / 2)
      }
    })
    context.stroke()
  }
}

/** Crops the drawing to its ink plus a small margin and returns it as a transparent PNG. */
function exportSignature(strokes: readonly Stroke[]): { dataUrl: string; width: number; height: number } | undefined {
  const points = strokes.flatMap(stroke => stroke.points)
  if (points.length < 2) return undefined
  const margin = 6
  const left = Math.max(0, Math.min(...points.map(point => point.x)) - margin)
  const top = Math.max(0, Math.min(...points.map(point => point.y)) - margin)
  const width = Math.min(padWidth, Math.max(...points.map(point => point.x)) + margin) - left
  const height = Math.min(padHeight, Math.max(...points.map(point => point.y)) + margin) - top
  const scale = 3
  const canvas = window.document.createElement('canvas')
  canvas.width = Math.ceil(width * scale)
  canvas.height = Math.ceil(height * scale)
  const context = canvas.getContext('2d')
  if (context === null) return undefined
  context.scale(scale, scale)
  context.translate(-left, -top)
  context.lineCap = 'round'
  context.lineJoin = 'round'
  context.lineWidth = 2.4
  for (const stroke of strokes) {
    context.strokeStyle = stroke.color
    context.beginPath()
    stroke.points.forEach((point, index) => {
      const previous = stroke.points[index - 1]
      if (previous === undefined) context.moveTo(point.x, point.y)
      else context.quadraticCurveTo(previous.x, previous.y, (previous.x + point.x) / 2, (previous.y + point.y) / 2)
    })
    context.stroke()
  }
  return { dataUrl: canvas.toDataURL('image/png'), width, height }
}

/** Draw a signature once with the trackpad or mouse; it is saved on this Mac for reuse. */
export function SignatureDialog() {
  const open = useToolStore(state => state.signatureDialogOpen)
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const [strokes, setStrokes] = useState<readonly Stroke[]>([])
  const [color, setColor] = useState<string>(inks[0].value)
  const drawing = useRef(false)

  useEffect(() => {
    const canvas = canvasRef.current
    if (canvas !== null) draw(canvas, strokes)
  }, [strokes])

  const close = (): void => {
    useToolStore.setState({ signatureDialogOpen: false })
    setStrokes([])
  }

  const pointOf = (event: ReactPointerEvent<HTMLCanvasElement>) => {
    const rect = event.currentTarget.getBoundingClientRect()
    return { x: event.clientX - rect.left, y: event.clientY - rect.top }
  }

  const save = (): void => {
    const image = exportSignature(strokes)
    if (image === undefined) return
    const id = newId()
    addSignature({ id, ...image })
    close()
    setTool('signature')
    useToolStore.setState({ armedSignatureId: id })
  }

  return (
    <Dialog.Root open={open} onOpenChange={next => { if (!next) close() }}>
      <Dialog.Portal>
        <Dialog.Backdrop className="pdf-dialog-backdrop" />
        <Dialog.Popup className="pdf-dialog pdf-signature-dialog">
          <Dialog.Title render={<h2 />}>Add signature</Dialog.Title>
          <Dialog.Description>Sign with your trackpad or mouse. It is saved on this Mac so you can place it on any PDF.</Dialog.Description>
          <canvas
            ref={canvas => {
              canvasRef.current = canvas
              if (canvas !== null && canvas.width !== padWidth * window.devicePixelRatio) {
                canvas.width = padWidth * window.devicePixelRatio
                canvas.height = padHeight * window.devicePixelRatio
                draw(canvas, strokes)
              }
            }}
            className="pdf-signature-pad"
            style={{ width: padWidth, height: padHeight }}
            onPointerDown={event => {
              event.currentTarget.setPointerCapture(event.pointerId)
              drawing.current = true
              setStrokes([...strokes, { color, points: [pointOf(event)] }])
            }}
            onPointerMove={event => {
              if (!drawing.current) return
              const point = pointOf(event)
              setStrokes(current => {
                const last = current.at(-1)
                return last === undefined ? current : [...current.slice(0, -1), { ...last, points: [...last.points, point] }]
              })
            }}
            onPointerUp={() => {
              drawing.current = false
            }}
          />
          <div className="pdf-signature-line" aria-hidden="true" />
          <div className="pdf-dialog-actions">
            <div className="segmented" role="group" aria-label="Ink colour">
              {inks.map(ink => (
                <button key={ink.value} type="button" aria-pressed={color === ink.value} onClick={() => setColor(ink.value)}>
                  <i className="pdf-ink-dot" style={{ background: ink.value }} />{ink.label}
                </button>
              ))}
            </div>
            <span className="toolbar-grow" />
            <button type="button" className="button is-quiet" disabled={strokes.length === 0} onClick={() => setStrokes([])}>Clear</button>
            <button type="button" className="button" onClick={close}>Cancel</button>
            <button type="button" className="button is-primary" disabled={strokes.length === 0} onClick={save}>Save signature</button>
          </div>
        </Dialog.Popup>
      </Dialog.Portal>
    </Dialog.Root>
  )
}
