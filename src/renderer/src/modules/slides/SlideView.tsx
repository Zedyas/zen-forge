import { memo, type CSSProperties, type ReactNode } from 'react'
import {
  cellInset,
  cornerRatio,
  holdsText,
  isLine,
  lineEnds,
  listIndent,
  listNumbers,
  polygonPoints,
  scalePath,
  type Paragraph,
  type ShapeElement,
  type Slide,
  type SlideElement,
  type TableElement,
  type TextBody,
  type TextRun,
} from './model'

/*
 * The one static drawing of a slide. Thumbnails, the editing canvas, the slideshow and printing all
 * use it, so a slide looks the same everywhere. The slide is laid out at one CSS pixel per point
 * and scaled as a whole, so text wraps identically at every size.
 */

/** PowerPoint's single line spacing is about 1.2 times the font size. */
export const singleLineHeight = 1.2

export function fontStack(font: string): string {
  return `"${font.replace(/"/g, '')}", Arial, Helvetica, sans-serif`
}

function runStyle(run: Omit<TextRun, 'text'>): CSSProperties {
  return {
    fontFamily: fontStack(run.font),
    fontSize: run.size,
    fontWeight: run.bold ? 700 : 400,
    fontStyle: run.italic ? 'italic' : 'normal',
    textDecorationLine: run.underline ? 'underline' : 'none',
    color: run.color,
    backgroundColor: run.highlight,
  }
}

/**
 * A paragraph takes its first run's font, so an empty line, its bullet and text typed into it have
 * that run's size and colour. Underline and highlight are left to the runs, which draw their own.
 */
function paragraphStyle(paragraph: Paragraph): CSSProperties {
  const first = paragraph.runs[0]
  return {
    ...(first === undefined ? {} : runStyle(first)),
    textDecorationLine: undefined,
    backgroundColor: undefined,
    textAlign: paragraph.align,
    lineHeight: singleLineHeight * paragraph.lineSpacing,
    paddingLeft: paragraph.list === 'none' ? 0 : listIndent * (paragraph.level + 1),
  }
}

function listClass(paragraph: Pick<Paragraph, 'list'>): string {
  return paragraph.list === 'bullet' ? ' is-bullet' : paragraph.list === 'number' ? ' is-number' : ''
}

/** Paragraphs and runs; the inline editor starts from exactly this markup and reads it back. */
export function TextParagraphs({ paragraphs }: { readonly paragraphs: readonly Paragraph[] }) {
  const numbers = listNumbers(paragraphs)
  return paragraphs.map((paragraph, index) => {
    const runs = paragraph.runs.filter(run => run.text !== '')
    return (
      <p
        key={index}
        className={`slide-paragraph${listClass(paragraph)}${runs.length === 0 ? ' is-empty' : ''}`}
        data-level={paragraph.level}
        data-number={numbers[index]}
        style={paragraphStyle(paragraph)}
      >
        {runs.length === 0 ? <br /> : runs.map((run, position) => <span key={position} style={runStyle(run)}>{run.text}</span>)}
      </p>
    )
  })
}

const justify = { top: 'flex-start', middle: 'center', bottom: 'flex-end' } as const

/** The text frame: insets and vertical alignment around the paragraphs, or around the editor. */
function TextFrame({ body, children }: { readonly body: TextBody; readonly children: ReactNode }) {
  const { inset } = body
  return (
    <div
      className="slide-text-frame"
      style={{ justifyContent: justify[body.verticalAlign], padding: `${inset.top}px ${inset.right}px ${inset.bottom}px ${inset.left}px` }}
    >
      {children}
    </div>
  )
}

function frameStyle(element: SlideElement): CSSProperties {
  return {
    left: element.x,
    top: element.y,
    width: element.width,
    height: element.height,
    transform: element.rotation === 0 ? undefined : `rotate(${element.rotation}deg)`,
  }
}

/** Arrowhead length and width for a line of `width` points, near PowerPoint's medium triangle. */
function arrowSize(width: number): number {
  return Math.max(6, width * 3)
}

function LineShape({ shape }: { readonly shape: ShapeElement }) {
  const [start, end] = lineEnds(shape)
  const stroke = shape.border?.color ?? 'transparent'
  const width = shape.border?.width ?? 1
  const a = { x: start.x - shape.x, y: start.y - shape.y }
  let b = { x: end.x - shape.x, y: end.y - shape.y }
  let head: string | undefined
  const length = Math.hypot(b.x - a.x, b.y - a.y)
  if (shape.geometry.type === 'arrow' && length > 0) {
    const size = Math.min(arrowSize(width), length)
    const ux = (b.x - a.x) / length
    const uy = (b.y - a.y) / length
    const base = { x: b.x - ux * size, y: b.y - uy * size }
    head = `${b.x},${b.y} ${base.x - (uy * size) / 2},${base.y + (ux * size) / 2} ${base.x + (uy * size) / 2},${base.y - (ux * size) / 2}`
    // The line stops inside the head, so its square end does not poke through the point.
    b = { x: base.x + ux * width * 0.5, y: base.y + uy * width * 0.5 }
  }
  return (
    <svg className="slide-geometry" width={Math.max(1, shape.width)} height={Math.max(1, shape.height)}>
      <line x1={a.x} y1={a.y} x2={b.x} y2={b.y} stroke={stroke} strokeWidth={width} />
      {head !== undefined && <polygon points={head} fill={stroke} />}
    </svg>
  )
}

function Outline({ element }: { readonly element: Exclude<SlideElement, { kind: 'image' | 'table' }> }) {
  const { width, height, fill, border } = element
  if (fill === undefined && border === undefined) return null
  const paint = { fill: fill ?? 'none', stroke: border?.color ?? 'none', strokeWidth: border?.width ?? 0 }
  const geometry = element.kind === 'shape' ? element.geometry : { type: 'rect' as const }
  let outline: ReactNode
  switch (geometry.type) {
    case 'roundRect': {
      const radius = Math.min(width, height) * cornerRatio
      outline = <rect width={width} height={height} rx={radius} ry={radius} {...paint} />
      break
    }
    case 'ellipse':
      outline = <ellipse cx={width / 2} cy={height / 2} rx={width / 2} ry={height / 2} {...paint} />
      break
    case 'triangle':
    case 'rightArrow':
      outline = <polygon points={polygonPoints(geometry.type, width, height).map(point => `${point.x},${point.y}`).join(' ')} strokeLinejoin="round" {...paint} />
      break
    case 'custom':
      outline = <path d={scalePath(geometry.path, geometry.pathWidth === 0 ? 1 : width / geometry.pathWidth, geometry.pathHeight === 0 ? 1 : height / geometry.pathHeight)} {...paint} />
      break
    default:
      outline = <rect width={width} height={height} {...paint} />
  }
  return <svg className="slide-geometry" width={Math.max(1, width)} height={Math.max(1, height)}>{outline}</svg>
}

/** Cells carry their row and column, so the canvas can find the cell under a click. */
function TableView({ table }: { readonly table: TableElement }) {
  const edge = table.border.width > 0 ? `${table.border.width}px solid ${table.border.color}` : 'none'
  return (
    <table className="slide-table" style={{ width: table.width, fontFamily: fontStack(table.font), fontSize: table.size, color: table.color }}>
      <colgroup>{table.columns.map((width, index) => <col key={index} style={{ width }} />)}</colgroup>
      <tbody>
        {table.rows.map((row, rowIndex) => (
          <tr key={rowIndex} style={{ height: row.height }}>
            {row.cells.map((cell, column) => (
              <td
                key={column}
                data-row={rowIndex}
                data-column={column}
                style={{
                  padding: `${cellInset.top}px ${cellInset.right}px ${cellInset.bottom}px ${cellInset.left}px`,
                  border: edge,
                  background: cell.fill,
                  color: cell.color,
                  fontWeight: cell.bold ? 700 : 400,
                  textAlign: cell.align,
                  verticalAlign: cell.verticalAlign,
                }}
              >
                {cell.text}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  )
}

interface ElementViewProps {
  readonly element: SlideElement
  /** Replaces the element's text, for the inline editor. */
  readonly text?: ReactNode
}

export function ElementView({ element, text }: ElementViewProps) {
  const frame = { className: 'slide-element', style: frameStyle(element), 'data-element-id': element.id }
  switch (element.kind) {
    case 'image':
      return <div {...frame}><img className="slide-image" src={element.src} alt="" draggable={false} /></div>
    case 'table':
      return <div {...frame}><TableView table={element} /></div>
    default:
      break
  }
  if (isLine(element)) return <div {...frame}><LineShape shape={element} /></div>
  const body = holdsText(element) ? element : undefined
  return (
    <div {...frame}>
      <Outline element={element} />
      {body !== undefined && (text !== undefined || body.paragraphs.length > 0) && (
        <TextFrame body={body}>
          {text ?? <div className="slide-text"><TextParagraphs paragraphs={body.paragraphs} /></div>}
        </TextFrame>
      )}
    </div>
  )
}

interface SlideViewProps {
  readonly slide: Slide
  /** The presentation's slide size in points. */
  readonly width: number
  readonly height: number
  /** CSS pixels per point. */
  readonly scale: number
  /** Left out, because something else draws it (the text editor). */
  readonly hiddenId?: string
}

/** A slide drawn at any scale. Memoised: a slide object that did not change is not drawn again. */
export const SlideView = memo(function SlideView({ slide, width, height, scale, hiddenId }: SlideViewProps) {
  return (
    <div className="slide-view" style={{ width: width * scale, height: height * scale }}>
      <div
        className="slide-surface"
        style={{
          width,
          height,
          // PowerPoint shows white through a background picture's transparent parts, not the colour.
          backgroundColor: slide.backgroundImage === undefined ? slide.background : '#ffffff',
          backgroundImage: slide.backgroundImage === undefined ? undefined : `url("${slide.backgroundImage}")`,
          transform: `scale(${scale})`,
        }}
      >
        {slide.elements.map(element => element.id === hiddenId ? null : <ElementView key={element.id} element={element} />)}
      </div>
    </div>
  )
})
