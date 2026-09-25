import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { documentPage, documentTheme, type BlockStyle } from './theme'

const width = 220

/** Draws the first page's text as a small PNG for the Home recents grid, in the document's own page and styles. */
export function renderPreview(doc: ProseMirrorNode): string {
  const page = documentPage(doc.attrs['page'])
  const theme = documentTheme(doc.attrs['theme'])
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

  // Pixels per point on the preview; page sizes are in twips (20 per point).
  const unit = width / (page.width / 20)
  const left = (page.marginLeft / 20) * unit
  const column = width - left - (page.marginRight / 20) * unit
  const bottom = height - (page.marginBottom / 20) * unit
  let y = (page.marginTop / 20) * unit

  const drawText = (text: string, style: BlockStyle): void => {
    context.font = `${style.italic ? 'italic ' : ''}${style.bold ? '700 ' : ''}${style.fontSize * unit}px "${style.fontFamily}", sans-serif`
    context.fillStyle = style.color
    const lineHeight = style.fontSize * unit * 1.15 * style.lineHeight
    y += style.spaceBefore * unit
    let line = ''
    const flush = (): void => {
      y += lineHeight
      context.fillText(line, left + style.indentLeft * unit, y)
      line = ''
    }
    for (const word of text.split(/\s+/)) {
      const next = line === '' ? word : `${line} ${word}`
      if (line !== '' && context.measureText(next).width > column - style.indentLeft * unit) flush()
      line = line === '' ? word : `${line} ${word}`
    }
    if (line !== '' || text === '') flush()
    y += style.spaceAfter * unit
  }

  // Walks the blocks in order and stops at the first page break or the bottom of the page.
  doc.descendants(node => {
    if (y > bottom || node.type.name === 'pageBreak') {
      y = height
      return false
    }
    const name = node.type.name
    if (name === 'heading') {
      const level: unknown = node.attrs['level']
      drawText(node.textContent, level === 2 ? theme.heading2 : level === 3 ? theme.heading3 : theme.heading1)
    } else if (name === 'title') drawText(node.textContent, theme.title)
    else if (name === 'paragraph') drawText(node.textContent, theme.normal)
    else if (name === 'tableRow') {
      const cells: string[] = []
      node.forEach(cell => cells.push(cell.textContent))
      drawText(cells.join('    '), { ...theme.normal, spaceAfter: 2 })
      return false
    } else if (name === 'horizontalRule') {
      context.fillStyle = '#bfbfbf'
      context.fillRect(left, y + 4 * unit, column, 1)
      y += 12 * unit
    }
    // Paragraphs inside lists and quotes are drawn as plain paragraphs, table rows as one line each.
    return !node.isTextblock
  })
  return canvas.toDataURL('image/png')
}
