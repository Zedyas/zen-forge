import { getMiddleCenterBias, GridCellKind, type DrawCellCallback } from '@glideapps/glide-data-grid'
import { cellFontSize } from './grid-theme'
import type { CellStyle } from './model/Workbook'
import { defaultFontSize, defaultRowHeight } from './model/workbook-data'

/*
 * Cell text that Glide cannot draw on its own: wrapped onto several lines, underlined or struck
 * through. Drawing and row fitting share `wrapLines`, so a fitted row is exactly as tall as its text.
 */

/** Space above and below the text block of a wrapped cell. */
const verticalPadding = 4
const lineSpacing = 1.3
const horizontalPadding = 8

/** Excel's default 11 pt shows at Ledger's standard 13 px, and every other size in proportion. */
export function fontPixels(style: CellStyle): number {
  return (style.fontSize ?? defaultFontSize) * cellFontSize / defaultFontSize
}

/** Glide's font style for a cell (weight, slant and size; the family is separate). */
export function fontStyle(style: CellStyle): string {
  return `${style.italic === true ? 'italic ' : ''}${style.bold === true ? '600 ' : ''}${fontPixels(style)}px`
}

/** A named family first, then the grid's own, for a Mac that lacks the named one. */
export function fontStack(family: string, fallback: string): string {
  return `"${family.replaceAll('"', '')}", ${fallback}`
}

/** Text split into lines no wider than `width`: at spaces, inside a word only when it is wider than a line. */
export function wrapLines(ctx: CanvasRenderingContext2D, text: string, width: number): readonly string[] {
  const lines: string[] = []
  const fits = (line: string): boolean => ctx.measureText(line).width <= width
  for (const paragraph of text.split(/\r?\n/)) {
    let line = ''
    for (const word of paragraph.split(/(?<=\s)/)) {
      if (fits(line + word.trimEnd()) || line === '') {
        line += word
      } else {
        lines.push(line.trimEnd())
        line = word
      }
      while (!fits(line.trimEnd()) && line.length > 1) {
        let cut = line.length - 1
        while (cut > 1 && !fits(line.slice(0, cut))) cut -= 1
        lines.push(line.slice(0, cut))
        line = line.slice(cut)
      }
    }
    lines.push(line.trimEnd())
  }
  return lines
}

/** Whether Ledger draws the cell's text itself instead of leaving it to Glide. */
export function drawsOwnText(style: CellStyle): boolean {
  return style.wrap === true || style.underline === true || style.strike === true
}

/** Draws a wrapped, underlined or struck-through cell's text, clipped to the cell. */
export function drawCellText(args: Parameters<DrawCellCallback>[0], text: string, style: CellStyle): void {
  const { ctx, rect, theme, cell } = args
  if (text === '') return
  const font = `${theme.baseFontStyle} ${theme.fontFamily}`
  const size = fontPixels(style)
  const lineHeight = size * lineSpacing
  ctx.save()
  ctx.beginPath()
  ctx.rect(rect.x, rect.y, rect.width, rect.height)
  ctx.clip()
  ctx.font = font
  ctx.fillStyle = theme.textDark
  ctx.strokeStyle = theme.textDark
  ctx.lineWidth = Math.max(1, Math.round(size / 13))
  ctx.textBaseline = 'middle'
  const align = cell.kind === GridCellKind.Text ? cell.contentAlign ?? 'left' : 'left'
  ctx.textAlign = align
  const inner = rect.width - horizontalPadding * 2
  const lines = style.wrap === true ? wrapLines(ctx, text, inner) : [text.split(/\r?\n/, 1)[0] ?? '']
  const blockHeight = lines.length * lineHeight
  const bias = getMiddleCenterBias(ctx, font)
  let middle = (blockHeight + verticalPadding * 2 > rect.height ? rect.y + verticalPadding : rect.y + (rect.height - blockHeight) / 2) + lineHeight / 2
  const x = align === 'right' ? rect.x + rect.width - horizontalPadding : align === 'center' ? rect.x + rect.width / 2 : rect.x + horizontalPadding
  for (const line of lines) {
    if (middle - lineHeight / 2 > rect.y + rect.height) break
    ctx.fillText(line, x, middle + bias)
    const width = Math.min(ctx.measureText(line).width, inner)
    const start = align === 'right' ? x - width : align === 'center' ? x - width / 2 : x
    const decorations = [style.underline === true ? size * 0.5 : undefined, style.strike === true ? size * 0.05 : undefined]
    decorations.forEach(offset => {
      if (offset === undefined || line === '') return
      const y = Math.round(middle + offset) + 0.5
      ctx.beginPath()
      ctx.moveTo(start, y)
      ctx.lineTo(start + width, y)
      ctx.stroke()
    })
    middle += lineHeight
  }
  ctx.restore()
}

let measuringContext: CanvasRenderingContext2D | null | undefined

/**
 * The row height a cell's text needs at a column width, in pixels; the workbook fits rows with it.
 * Only wrapped text and fonts larger than the standard size need more than the default row.
 */
export function measureCellHeight(text: string, style: CellStyle, width: number): number {
  if (style.wrap !== true && fontPixels(style) <= cellFontSize) return defaultRowHeight
  measuringContext ??= document.createElement('canvas').getContext('2d')
  if (measuringContext === null) return defaultRowHeight
  const family = getComputedStyle(document.documentElement).getPropertyValue('--font-ui').trim()
  measuringContext.font = `${fontStyle(style)} ${style.fontFamily === undefined ? family : fontStack(style.fontFamily, family)}`
  const lines = style.wrap === true ? wrapLines(measuringContext, text, width - horizontalPadding * 2).length : 1
  return Math.max(defaultRowHeight, Math.ceil(lines * fontPixels(style) * lineSpacing + verticalPadding * 2))
}
