import { formatCellValue } from './model/format'
import type { Workbook } from './model/Workbook'

const columns = 5
const rows = 10
const cellWidth = 64
const cellHeight = 18

/** Draws the top-left cells of a sheet as a small PNG for the Home recents grid. */
export function renderPreview(workbook: Workbook, sheetId: number): string {
  const scale = 2
  const canvas = document.createElement('canvas')
  canvas.width = columns * cellWidth * scale
  canvas.height = rows * cellHeight * scale
  const context = canvas.getContext('2d')
  if (context === null) return ''
  context.scale(scale, scale)
  context.fillStyle = '#ffffff'
  context.fillRect(0, 0, columns * cellWidth, rows * cellHeight)
  for (let row = 0; row < rows; row += 1) {
    for (let col = 0; col < columns; col += 1) {
      const style = workbook.getStyle(sheetId, [col, row])
      const value = workbook.getCellValue(sheetId, [col, row])
      const x = col * cellWidth
      const y = row * cellHeight
      if (style.fillColor !== undefined) {
        context.fillStyle = style.fillColor
        context.fillRect(x, y, cellWidth, cellHeight)
      }
      context.strokeStyle = '#e3e5e8'
      context.strokeRect(x + 0.5, y + 0.5, cellWidth, cellHeight)
      const text = formatCellValue(value, style)
      if (text === '') continue
      context.fillStyle = style.textColor ?? '#1d1f23'
      context.font = `${style.bold === true ? '600 ' : ''}10px system-ui, -apple-system, sans-serif`
      context.textBaseline = 'middle'
      const numeric = typeof value === 'number'
      context.textAlign = numeric ? 'right' : 'left'
      context.save()
      context.beginPath()
      context.rect(x, y, cellWidth, cellHeight)
      context.clip()
      context.fillText(text, numeric ? x + cellWidth - 4 : x + 4, y + cellHeight / 2)
      context.restore()
    }
  }
  return canvas.toDataURL('image/png')
}
