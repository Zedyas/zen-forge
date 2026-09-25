// @vitest-environment happy-dom
import PptxGenJS from 'pptxgenjs'
import { describe, expect, it } from 'vitest'
import { createImportReport } from '@shared/fidelity'
import { newId, paragraphOf, textRun, type Presentation, type SlideElement } from './model'
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
    { runs: [textRun('Revenue '), textRun('up 12%', { bold: true })], align: 'left', bullet: true, level: 0 },
    { runs: [textRun('Costs flat', { italic: true, color: '#c9352b' })], align: 'left', bullet: true, level: 1 },
  ],
}
const shape: SlideElement = {
  kind: 'shape', id: newId(), x: 600, y: 150, width: 200, height: 120, rotation: 15, geometry: { type: 'roundRect' },
  fill: '#4a78c2', border: { color: '#222222', width: 2 }, flipH: false, flipV: false,
  paragraphs: [paragraphOf('Goal', {}, { align: 'center' })], verticalAlign: 'middle', inset: frame.inset,
}
const picture: SlideElement = { kind: 'image', id: newId(), x: 620, y: 320, width: 160, height: 160, rotation: 0, src: png }

const presentation: Presentation = {
  width: 960,
  height: 540,
  slides: [{ id: newId(), background: '#fdf6e3', elements: [title, bullets, shape, picture], notes: 'Speak slowly' }],
}

describe('.pptx round trip', () => {
  it('keeps text, bullets, bold, shapes, pictures, positions and notes, and reports nothing lost', async () => {
    const { presentation: reopened, findings } = await readPptx(await writePptx(presentation))
    const slide = reopened.slides[0]

    expect(findings).toEqual([])
    expect(reopened).toMatchObject({ width: 960, height: 540 })
    expect(slide).toMatchObject({ background: '#fdf6e3', notes: 'Speak slowly' })
    expect(slide?.elements.map(element => element.kind)).toEqual(['text', 'text', 'shape', 'image'])

    expect(slide?.elements[1]).toMatchObject({
      paragraphs: [
        { bullet: true, level: 0, runs: [{ text: 'Revenue ', bold: false }, { text: 'up 12%', bold: true, size: 24, font: 'Arial' }] },
        { bullet: true, level: 1, runs: [{ text: 'Costs flat', italic: true, color: '#c9352b' }] },
      ],
    })
    expect(slide?.elements[0]).toMatchObject({ paragraphs: [{ bullet: false, runs: [{ text: 'Quarterly review', size: 40 }] }] })
    expect(slide?.elements[2]).toMatchObject({
      geometry: { type: 'roundRect' }, fill: '#4a78c2', border: { color: '#222222', width: 2 }, rotation: 15,
      verticalAlign: 'middle', paragraphs: [{ align: 'center', runs: [{ text: 'Goal' }] }],
    })
    expect(slide?.elements[3]).toMatchObject({ kind: 'image', src: png })

    presentation.slides[0]?.elements.forEach((original, index) => {
      const copy = slide?.elements[index]
      for (const key of ['x', 'y', 'width', 'height'] as const) expect(copy?.[key]).toBeCloseTo(original[key], 0)
    })
  })
})

describe('.pptx import report', () => {
  it('reports charts and tables as lost on save', async () => {
    const pptx = new PptxGenJS()
    const slide = pptx.addSlide()
    slide.addText('Results', { x: 0.5, y: 0.3, w: 6, h: 0.8 })
    slide.addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [4, 7] }], { x: 0.5, y: 1.2, w: 4, h: 3 })
    slide.addTable([[{ text: 'Region' }, { text: 'Total' }], [{ text: 'North' }, { text: '12' }]], { x: 5, y: 1.2, w: 4 })
    const written = await pptx.write({ outputType: 'uint8array' })
    if (!(written instanceof Uint8Array)) throw new Error('pptxgenjs did not write bytes')

    const { presentation: imported, findings } = await readPptx(written)

    expect(findings).toEqual(expect.arrayContaining([
      expect.objectContaining({ construct: 'Charts', severity: 'dropped', location: 'Slide 1' }),
      expect.objectContaining({ construct: 'Tables', severity: 'dropped', location: 'Slide 1' }),
    ]))
    expect(createImportReport('Results.pptx', findings).severity).toBe('dropped')
    expect(imported.slides[0]?.elements).toHaveLength(1)
  })
})
