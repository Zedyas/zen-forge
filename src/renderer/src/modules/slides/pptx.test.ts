// @vitest-environment happy-dom
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import PptxGenJS from 'pptxgenjs'
import { describe, expect, it } from 'vitest'
import { createImportReport } from '@shared/fidelity'
import { createSlide, newId, newTable, paragraphOf, textRun, updateCell, type Paragraph, type Presentation, type SlideElement } from './model'
import { writePptx } from './pptx-export'
import { readPptx } from './pptx-import'

/** A 2 × 2 red PNG. */
const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFklEQVR4nGP8z8DAwMDAxMDAwMDAAAANHQEDasKb6QAAAABJRU5ErkJggg=='

const frame = { rotation: 0, verticalAlign: 'top' as const, inset: { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 } }

const title: SlideElement = {
  kind: 'text', id: newId(), x: 58, y: 40, width: 844, height: 80, ...frame,
  paragraphs: [paragraphOf('Quarterly review', { size: 40 })],
}
const bullets: SlideElement = {
  kind: 'text', id: newId(), x: 58, y: 140, width: 500, height: 300, ...frame,
  paragraphs: [
    { runs: [textRun('Revenue '), textRun('up 12%', { bold: true, highlight: '#ffff00' })], align: 'left', list: 'bullet', level: 0, lineSpacing: 1.5 },
    { runs: [textRun('Costs flat', { italic: true, color: '#c9352b' })], align: 'left', list: 'bullet', level: 1, lineSpacing: 1 },
    { runs: [textRun('First step')], align: 'left', list: 'number', level: 0, lineSpacing: 1 },
  ],
}
const shape: SlideElement = {
  kind: 'shape', id: newId(), x: 600, y: 150, width: 200, height: 120, rotation: 15, geometry: { type: 'roundRect' },
  fill: '#4a78c2', border: { color: '#222222', width: 2 }, flipH: false, flipV: false,
  paragraphs: [paragraphOf('Goal', {}, { align: 'center' })], verticalAlign: 'middle', inset: frame.inset,
}
const picture: SlideElement = { kind: 'image', id: newId(), x: 620, y: 320, width: 160, height: 160, rotation: 0, src: png }
const table = updateCell(updateCell(newTable(3, 2, { width: 960, height: 540, theme: 'light' }), 0, 0, { text: 'Region' }), 1, 1, { text: '12', align: 'right', fill: '#dbe8fb' })
// Empty layout boxes in a dark theme (light text), and a box with a blank line between two lines and a tab.
const layout = createSlide('titleContent', { width: 960, height: 540, theme: 'slate' })
const spaced: SlideElement = {
  kind: 'text', id: newId(), x: 58, y: 300, width: 500, height: 100, ...frame,
  paragraphs: [paragraphOf('Name\tValue'), paragraphOf('', { size: 32, color: '#c9352b', bold: true, font: 'Georgia' }), paragraphOf('Below')],
}

const presentation: Presentation = {
  width: 960,
  height: 540,
  theme: 'light',
  slides: [
    { id: newId(), background: '#fdf6e3', elements: [title, bullets, shape, picture], notes: 'Speak slowly' },
    { id: newId(), background: '#ffffff', backgroundImage: png, elements: [table], notes: '' },
    { ...layout, elements: [...layout.elements, spaced] },
  ],
}

/** What an empty line keeps: its look, not its (absent) text. */
function look(paragraph: Paragraph | undefined) {
  const run = paragraph?.runs[0]
  return { align: paragraph?.align, list: paragraph?.list, size: run?.size, color: run?.color, font: run?.font, bold: run?.bold }
}

describe('.pptx round trip', () => {
  it('keeps text, lists, bold, highlight, shapes, pictures, tables, positions and notes, and reports nothing lost', async () => {
    const { presentation: reopened, findings } = await readPptx(await writePptx(presentation))
    const [first, second, third] = reopened.slides

    expect(findings).toEqual([])
    expect(reopened).toMatchObject({ width: 960, height: 540 })
    expect(first).toMatchObject({ background: '#fdf6e3', notes: 'Speak slowly' })
    expect(first?.elements.map(element => element.kind)).toEqual(['text', 'text', 'shape', 'image'])
    expect(first?.elements[1]).toMatchObject({
      paragraphs: [
        { list: 'bullet', level: 0, lineSpacing: 1.5, runs: [{ text: 'Revenue ', bold: false }, { text: 'up 12%', bold: true, highlight: '#ffff00', size: 24, font: 'Arial' }] },
        { list: 'bullet', level: 1, runs: [{ text: 'Costs flat', italic: true, color: '#c9352b' }] },
        { list: 'number', level: 0, runs: [{ text: 'First step' }] },
      ],
    })
    expect(first?.elements[0]).toMatchObject({ paragraphs: [{ list: 'none', runs: [{ text: 'Quarterly review', size: 40 }] }] })
    expect(first?.elements[2]).toMatchObject({
      geometry: { type: 'roundRect' }, fill: '#4a78c2', border: { color: '#222222', width: 2 }, rotation: 15,
      verticalAlign: 'middle', paragraphs: [{ align: 'center', runs: [{ text: 'Goal' }] }],
    })
    expect(first?.elements[3]).toMatchObject({ kind: 'image', src: png })
    presentation.slides[0]?.elements.forEach((original, index) => {
      const copy = first?.elements[index]
      for (const key of ['x', 'y', 'width', 'height'] as const) expect(copy?.[key]).toBeCloseTo(original[key], 0)
    })

    expect(second?.backgroundImage).toBe(png)
    const reopenedTable = second?.elements[0]
    expect(reopenedTable).toMatchObject({ kind: 'table', size: 16, font: 'Arial' })
    if (reopenedTable?.kind !== 'table') throw new Error('The table did not come back')
    expect(reopenedTable.columns).toHaveLength(2)
    expect(reopenedTable.rows.map(row => row.cells.map(cell => cell.text))).toEqual([['Region', ''], ['', '12'], ['', '']])
    expect(reopenedTable.rows[0]?.cells[0]).toMatchObject({ bold: true, fill: '#4a78c2', color: '#ffffff' })
    // An empty header cell keeps its bold white text, so the header row stays on.
    expect(reopenedTable.rows[0]?.cells[1]).toMatchObject({ text: '', bold: true, fill: '#4a78c2', color: '#ffffff' })
    expect(reopenedTable.rows[1]?.cells[1]).toMatchObject({ align: 'right', fill: '#dbe8fb' })
    expect(reopenedTable.x).toBeCloseTo(table.x, 0)
    expect(reopenedTable.width).toBeCloseTo(table.width, 0)

    const originalLooks = presentation.slides[2]?.elements.flatMap(element => element.kind === 'text' ? element.paragraphs.map(look) : [])
    expect(third?.elements.flatMap(element => element.kind === 'text' ? element.paragraphs.map(look) : [])).toEqual(originalLooks)
    const lines = third?.elements[2]
    expect(lines?.kind === 'text' ? lines.paragraphs.map(paragraph => paragraph.runs.map(run => run.text).join('')) : undefined).toEqual(['Name\tValue', '', 'Below'])
  })
})

describe('.pptx import report', () => {
  it('reports charts as lost on save, and keeps tables', async () => {
    const pptx = new PptxGenJS()
    const slide = pptx.addSlide()
    slide.addText('Results', { x: 0.5, y: 0.3, w: 6, h: 0.8 })
    slide.addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [4, 7] }], { x: 0.5, y: 1.2, w: 4, h: 3 })
    slide.addTable([[{ text: 'Region' }, { text: 'Total' }], [{ text: 'North' }, { text: '12' }]], { x: 5, y: 1.2, w: 4 })
    const written = await pptx.write({ outputType: 'uint8array' })
    if (!(written instanceof Uint8Array)) throw new Error('pptxgenjs did not write bytes')

    const { presentation: imported, findings } = await readPptx(written)

    expect(findings).toEqual([expect.objectContaining({ construct: 'Charts', severity: 'dropped', location: 'Slide 1', suggestedAlternative: expect.stringContaining('removed when saved') })])
    expect(createImportReport('Results.pptx', findings).severity).toBe('dropped')
    expect(imported.slides[0]?.elements.map(element => element.kind)).toEqual(['shape', 'table'])
  })
})

describe('.pptx import report for tables and text', () => {
  it('reports what a table cell or a text box would lose, so saving asks first', async () => {
    const pptx = new PptxGenJS()
    const slide = pptx.addSlide()
    const none = { type: 'none' as const }
    const line = { type: 'solid' as const, pt: 1, color: '000000' }
    slide.addTable([[
      { text: [{ text: 'Total: ' }, { text: '12', options: { bold: true } }] },
      { text: [{ text: 'one', options: { bullet: true, breakLine: true } }, { text: 'two', options: { bullet: true } }] },
      { text: 'edges', options: { border: [line, none, line, none] } },
    ]], { x: 0.5, y: 0.5, w: 6 })
    slide.addText('A title that must not wrap', { x: 0.5, y: 3, w: 2, h: 1, wrap: false })
    const written = await pptx.write({ outputType: 'uint8array' })
    if (!(written instanceof Uint8Array)) throw new Error('pptxgenjs did not write bytes')

    const { findings } = await readPptx(written)

    expect(findings.map(finding => [finding.construct, finding.severity])).toEqual(expect.arrayContaining([
      ['Mixed formatting inside a table cell', 'degraded'],
      ['Bullets and numbering in table cells', 'degraded'],
      ['Table borders that differ between cells', 'degraded'],
      ['Text set not to wrap', 'degraded'],
    ]))
  })

  it('gives a plain message for a file pptxtojson cannot read', async () => {
    const parts = unzipSync(await writePptx(presentation))
    parts['ppt/slides/slide1.xml'] = strToU8('<p:sld')
    await expect(readPptx(zipSync(parts))).rejects.toThrow('This presentation couldn\'t be read. It may be damaged, or use something Slides doesn\'t support.')
  })
})

describe('.pptx slide order', () => {
  it('follows the presentation\'s slide list, not the slide file numbers', async () => {
    const two: Presentation = {
      ...presentation,
      slides: [
        { id: newId(), background: '#ffffff', elements: [{ ...title, paragraphs: [paragraphOf('First')] }], notes: '' },
        { id: newId(), background: '#ffffff', elements: [{ ...title, paragraphs: [paragraphOf('Second')] }], notes: '' },
      ],
    }
    const parts = unzipSync(await writePptx(two))
    const xml = strFromU8(parts['ppt/presentation.xml'] ?? new Uint8Array())
    // Swap the two entries of the slide list, as reordering in PowerPoint can leave the files as they were.
    const [first, second] = xml.match(/<p:sldId\b[^>]*\/>/g) ?? []
    if (first === undefined || second === undefined) throw new Error('No slide list written')
    parts['ppt/presentation.xml'] = strToU8(xml.replace(first, '@first@').replace(second, first).replace('@first@', second))

    const { presentation: reopened } = await readPptx(zipSync(parts))

    expect(reopened.slides.map(slide => slide.elements[0]?.kind === 'text' ? slide.elements[0].paragraphs[0]?.runs[0]?.text : undefined)).toEqual(['Second', 'First'])
  })
})
