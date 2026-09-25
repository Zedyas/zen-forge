import { z } from 'zod'
import type { SlideElement } from './model'

/*
 * Copied slide objects travel through the system clipboard as JSON, which any app can write. A
 * paste only accepts JSON that has exactly the model's shape, with pictures limited to embedded
 * PNG, JPEG and GIF data, so nothing pasted can load from the network or break drawing.
 */

const color = z.string().regex(/^#[0-9a-f]{6}(?:[0-9a-f]{2})?$/i)
const length = z.number().finite().min(-100_000).max(100_000)
const size = z.number().finite().min(0).max(100_000)
const align = z.enum(['left', 'center', 'right', 'justify'])
const verticalAlign = z.enum(['top', 'middle', 'bottom'])
const border = z.object({ color, width: z.number().finite().min(0).max(1000) })
const inset = z.object({ top: size, right: size, bottom: size, left: size })

const run = z.object({
  text: z.string(),
  bold: z.boolean(),
  italic: z.boolean(),
  underline: z.boolean(),
  color,
  highlight: color.optional(),
  size: z.number().finite().positive().max(4000),
  font: z.string().max(200),
})

const paragraph = z.object({
  runs: z.array(run).min(1),
  align,
  list: z.enum(['none', 'bullet', 'number']),
  level: z.number().int().min(0).max(8),
  lineSpacing: z.number().finite().positive().max(10),
})

const frame = { id: z.string(), x: length, y: length, width: size, height: size, rotation: z.number().finite() }
const body = { paragraphs: z.array(paragraph), verticalAlign, inset }

const textBox = z.object({
  kind: z.literal('text'),
  ...frame,
  ...body,
  paragraphs: z.array(paragraph).min(1),
  fill: color.optional(),
  border: border.optional(),
  prompt: z.string().optional(),
})

const shape = z.object({
  kind: z.literal('shape'),
  ...frame,
  ...body,
  geometry: z.union([
    z.object({ type: z.enum(['rect', 'roundRect', 'ellipse', 'triangle', 'rightArrow', 'line', 'arrow']) }),
    z.object({ type: z.literal('custom'), path: z.string().max(200_000), pathWidth: size, pathHeight: size, preset: z.string().max(100).optional() }),
  ]),
  fill: color.optional(),
  border: border.optional(),
  flipH: z.boolean(),
  flipV: z.boolean(),
})

const image = z.object({
  kind: z.literal('image'),
  ...frame,
  src: z.string().regex(/^data:image\/(?:png|jpeg|gif);base64,[A-Za-z0-9+/]+=*$/),
})

const cell = z.object({ text: z.string(), bold: z.boolean(), color: color.optional(), fill: color.optional(), align, verticalAlign })

const table = z.object({
  kind: z.literal('table'),
  ...frame,
  columns: z.array(z.number().finite().positive()).min(1).max(100),
  rows: z.array(z.object({ height: size, cells: z.array(cell) })).min(1).max(1000),
  font: z.string().max(200),
  size: z.number().finite().positive().max(4000),
  color,
  border,
}).refine(value => value.rows.every(row => row.cells.length === value.columns.length), 'Every row needs one cell per column.')

const elements = z.array(z.discriminatedUnion('kind', [textBox, shape, image, table])).min(1)

/** The elements in pasted JSON, or undefined when it is not exactly a list of slide objects. */
export function pastedElements(json: string): SlideElement[] | undefined {
  try {
    const result = elements.safeParse(JSON.parse(json))
    return result.success ? result.data : undefined
  } catch {
    return undefined
  }
}
