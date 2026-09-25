import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { PageFormat } from './page'

const width = 220

/** Point sizes of the text blocks, as the page shows them. */
const fontSizes: Readonly<Record<string, number>> = { 1: 20, 2: 16, 3: 13 }

/** Draws the first page's text as a small PNG for the Home recents grid. */
export function renderPreview(doc: ProseMirrorNode, page: PageFormat): string {
  const scale = 2
  const height = Math.round(width * (page.height / page.width))
  const canvas = document.createElement('canvas')
  canvas.width = width * scale
  canvas.height = height * scale
  const context = canvas.getContext('2d')
  if (context === null) return ''
  context.scale(scale, scale)
  context.fillStyle = '#ffffff'
  context.fillRect(0, 0, width, height)

  // Pixels per point on the preview.
  const unit = width / (page.width / 20)
  const margin = (page.margin / 20) * unit
  const column = width - 2 * margin
  let y = margin

  const drawText = (text: string, size: number, bold: boolean, indent: number): void => {
    context.font = `${bold ? '700 ' : ''}${size * unit}px Arial, sans-serif`
    context.fillStyle = '#1d1f23'
    const lineHeight = size * unit * 1.15
    let line = ''
    const flush = (): void => {
      y += lineHeight
      context.fillText(line, margin + indent, y)
      line = ''
    }
    for (const word of text.split(/\s+/)) {
      const next = line === '' ? word : `${line} ${word}`
      if (line !== '' && context.measureText(next).width > column - indent) flush()
      line = line === '' ? word : `${line} ${word}`
    }
    if (line !== '' || text === '') flush()
    y += (bold ? 4 : 8) * unit
  }

  // Walks the blocks in order and stops at the first page break or the bottom of the page.
  doc.descendants(node => {
    if (y > height - margin || node.type.name === 'pageBreak') {
      y = height
      return false
    }
    if (node.type.name === 'heading') drawText(node.textContent, fontSizes[String(node.attrs['level'])] ?? 13, true, 0)
    else if (node.type.name === 'paragraph') drawText(node.textContent, 11, false, 0)
    else if (node.type.name === 'horizontalRule') {
      context.fillStyle = '#bfbfbf'
      context.fillRect(margin, y + 4 * unit, column, 1)
      y += 12 * unit
    }
    // Paragraphs inside lists, quotes and tables are drawn as plain paragraphs.
    return !node.isTextblock
  })
  return canvas.toDataURL('image/png')
}
