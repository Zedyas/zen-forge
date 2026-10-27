// @vitest-environment happy-dom
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import PptxGenJS from 'pptxgenjs'
import { describe, expect, it } from 'vitest'
import { createImportReport } from '@shared/fidelity'
import { createSlide, lineEnds, newId, newTable, paragraphOf, textRun, updateCell, withHeaderRow, type Paragraph, type Presentation, type SlideElement } from './model'
import { writePptx } from './pptx-export'
import { readPptx } from './pptx-import'
import { findTheme } from './themes'

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

describe('.pptx table borders', () => {
  it('saves a table whose border is invisible without visible borders', async () => {
    const invisible = { ...newTable(2, 2, { width: 960, height: 540, theme: 'light' }), border: { color: '#00000000', width: 1.5 } }
    const deck: Presentation = { ...presentation, slides: [{ id: newId(), background: '#ffffff', elements: [invisible], notes: '' }] }

    const { presentation: reopened } = await readPptx(await writePptx(deck))

    const table = reopened.slides[0]?.elements[0]
    expect(table).toMatchObject({ kind: 'table', border: { width: 0 } })
  })
})

describe('.pptx font names', () => {
  it('saves a font name with &, < or " as a file that reopens, with the same name', async () => {
    const font = 'Fira & Co <Sans>'
    const box: SlideElement = { ...title, paragraphs: [paragraphOf('Named', { font }), paragraphOf('', { font }), paragraphOf('Quoted', { font: 'The "Font"' })] }
    const table = { ...newTable(1, 1, { width: 960, height: 540, theme: 'light' }), font }
    const deck: Presentation = { ...presentation, slides: [{ id: newId(), background: '#ffffff', elements: [box, table], notes: '' }] }

    const bytes = await writePptx(deck)
    const { presentation: reopened } = await readPptx(bytes)

    // PowerPoint refuses a slide whose XML is not well formed; pptxtojson reads past it.
    const xml = strFromU8(unzipSync(bytes)['ppt/slides/slide1.xml'] ?? new Uint8Array())
    expect(new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagName('parsererror')).toHaveLength(0)
    const [text, grid] = reopened.slides[0]?.elements ?? []
    expect(text?.kind === 'text' ? text.paragraphs.slice(0, 2).map(paragraph => paragraph.runs[0]?.font) : undefined).toEqual([font, font])
    expect(grid).toMatchObject({ kind: 'table', font })
  })
})

describe('.pptx control characters', () => {
  it('saves pasted vertical tabs as line breaks in text and cells and as new lines in notes, and drops other control characters', async () => {
    const pasted = 'one\u000btwo\u0001'
    const box: SlideElement = { ...title, paragraphs: [paragraphOf(pasted)] }
    const table = updateCell(newTable(1, 1, { width: 960, height: 540, theme: 'light' }), 0, 0, { text: pasted })
    const deck: Presentation = { ...presentation, slides: [{ id: newId(), background: '#ffffff', elements: [box, table], notes: pasted }] }

    const parts = unzipSync(await writePptx(deck))

    const slideXml = strFromU8(parts['ppt/slides/slide1.xml'] ?? new Uint8Array())
    // One line break in the text box and one in the cell, each between the two words.
    expect(slideXml.match(/<a:t>one<\/a:t><\/a:r><a:br\/><a:r>.*?<a:t>two<\/a:t>/g)).toHaveLength(2)
    const notesXml = strFromU8(parts['ppt/notesSlides/notesSlide1.xml'] ?? new Uint8Array())
    expect(notesXml).toMatch(/<a:t>one\r?\ntwo<\/a:t>/)
    for (const xml of [slideXml, notesXml]) {
      expect(xml).not.toContain('onetwo')
      expect(xml).not.toContain('\u000b')
      expect(xml).not.toContain('\u0001')
    }
  })
})

describe('.pptx equations', () => {
  it('places the picture of an equation in a placeholder where its layout puts the placeholder', async () => {
    const deck: Presentation = { ...presentation, slides: [{ id: newId(), background: '#ffffff', elements: [picture], notes: '' }] }
    const parts = unzipSync(await writePptx(deck))
    const slideXml = strFromU8(parts['ppt/slides/slide1.xml'] ?? new Uint8Array())
    // The equation's fallback picture reuses the picture's image part.
    const image = /r:embed="([^"]+)"/.exec(slideXml)?.[1] ?? ''
    const shape = (content: string): string => `<p:sp><p:nvSpPr><p:cNvPr id="50" name="Equation"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>${content}</p:sp>`
    const equation = '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">'
      + '<mc:Choice Requires="a14">' + shape('<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a14:m><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara></a14:m></a:p></p:txBody>') + '</mc:Choice>'
      + `<mc:Fallback>${shape(`<p:spPr><a:blipFill><a:blip r:embed="${image}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr>`)}</mc:Fallback></mc:AlternateContent>`
    parts['ppt/slides/slide1.xml'] = strToU8(slideXml.replace('</p:spTree>', `${equation}</p:spTree>`))
    // The layout's placeholder sits 1 inch in from the top-left corner, 2 inches wide and 1 inch high.
    parts['ppt/slideLayouts/slideLayout1.xml'] = strToU8(strFromU8(parts['ppt/slideLayouts/slideLayout1.xml'] ?? new Uint8Array()).replace('</p:spTree>',
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>'
      + '<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></p:spTree>'))

    const { presentation: reopened, findings } = await readPptx(zipSync(parts))

    expect(findings.map(finding => finding.construct)).toContain('Equations')
    expect(reopened.slides[0]?.elements[1]).toMatchObject({ kind: 'image', x: 72, y: 72, width: 144, height: 72 })
  })
})

describe('.pptx flipped groups', () => {
  it('mirrors where a flipped group\'s shapes and lines sit, and reports the shapes as flipped', async () => {
    const parts = unzipSync(await writePptx({ ...presentation, slides: [{ id: newId(), background: '#ffffff', elements: [], notes: '' }] }))
    // A group 2 × 1 inches, flipped left to right, holding a half-inch square at its left end and a line falling to the right.
    const child = (id: number, geometry: string, x: number, cx: number, cy: number): string => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Child ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
      + `<p:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${geometry}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr></p:sp>`
    const group = '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="60" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm flipH="1"><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/>'
      + `<a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></p:grpSpPr>${child(61, 'rect', 0, 457200, 457200)}${child(62, 'line', 914400, 914400, 914400)}</p:grpSp>`
    parts['ppt/slides/slide1.xml'] = strToU8(strFromU8(parts['ppt/slides/slide1.xml'] ?? new Uint8Array()).replace('</p:spTree>', `${group}</p:spTree>`))

    const { presentation: reopened, findings } = await readPptx(zipSync(parts))

    const [square, line] = reopened.slides[0]?.elements ?? []
    expect(square).toMatchObject({ kind: 'shape', x: 108, y: 0, width: 36, height: 36 })
    // The line fell to the right from its top-left corner; mirrored, it falls to the left from the top-right of the group's left half.
    expect(line?.kind === 'shape' ? lineEnds(line) : undefined).toEqual([{ x: 72, y: 0 }, { x: 0, y: 72 }])
    expect(findings.map(finding => finding.construct)).toContain('Flipped shapes and pictures')
  })
})

describe('.pptx nested groups', () => {
  it('mirrors a shape once for the flipped group that holds its unflipped group', async () => {
    const parts = unzipSync(await writePptx({ ...presentation, slides: [{ id: newId(), background: '#ffffff', elements: [], notes: '' }] }))
    const square = '<p:sp><p:nvSpPr><p:cNvPr id="63" name="Square"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="228600" cy="228600"/></a:xfrm>'
      + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>'
    const group = (id: number, flip: string, size: number, content: string): string => `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`
      + `<p:grpSpPr><a:xfrm${flip}><a:off x="0" y="0"/><a:ext cx="${size}" cy="${size}"/><a:chOff x="0" y="0"/><a:chExt cx="${size}" cy="${size}"/></a:xfrm></p:grpSpPr>${content}</p:grpSp>`
    // A 144 pt group flipped left to right holds an unflipped 72 pt group holding an 18 pt square at its left edge.
    const nested = group(60, ' flipH="1"', 1828800, group(61, '', 914400, square))
    parts['ppt/slides/slide1.xml'] = strToU8(strFromU8(parts['ppt/slides/slide1.xml'] ?? new Uint8Array()).replace('</p:spTree>', `${nested}</p:spTree>`))

    const { presentation: reopened } = await readPptx(zipSync(parts))

    expect(reopened.slides[0]?.elements[0]).toMatchObject({ kind: 'shape', x: 126, y: 0, width: 18, height: 18 })
  })
})

describe('.pptx placeholder bullets', () => {
  it('leaves out the bullets a content placeholder\'s layout turns off, and keeps them where it does not', async () => {
    const box = (text: string, y: number): SlideElement => ({ ...title, id: newId(), y, paragraphs: [paragraphOf(text)] })
    const deck: Presentation = { ...presentation, slides: [{ id: newId(), background: '#ffffff', elements: [box('Plain', 40), box('Bulleted', 200)], notes: '' }] }
    const parts = unzipSync(await writePptx(deck))
    // Make both boxes content placeholders that state no bullets of their own; the layout's idx 1 turns bullets off, its idx 2 does not.
    let index = 0
    parts['ppt/slides/slide1.xml'] = strToU8(strFromU8(parts['ppt/slides/slide1.xml'] ?? new Uint8Array())
      .replace(/<p:nvPr><\/p:nvPr>/g, () => `<p:nvPr><p:ph idx="${++index}"/></p:nvPr>`)
      .replace(/<a:buNone\/>/g, ''))
    const layoutPlaceholder = (idx: number, levels: string): string => `<p:sp><p:nvSpPr><p:cNvPr id="${idx + 1}" name="Body ${idx}"/><p:cNvSpPr/><p:nvPr><p:ph idx="${idx}"/></p:nvPr></p:nvSpPr>`
      + `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle>${levels}</a:lstStyle><a:p/></p:txBody></p:sp>`
    parts['ppt/slideLayouts/slideLayout1.xml'] = strToU8(strFromU8(parts['ppt/slideLayouts/slideLayout1.xml'] ?? new Uint8Array())
      .replace('</p:spTree>', `${layoutPlaceholder(1, '<a:lvl1pPr><a:buNone/></a:lvl1pPr>')}${layoutPlaceholder(2, '')}</p:spTree>`))

    const { presentation: reopened } = await readPptx(zipSync(parts))

    const lists = reopened.slides[0]?.elements.map(element => element.kind === 'text' ? [element.paragraphs[0]?.runs[0]?.text, element.paragraphs[0]?.list] : [])
    expect(lists).toEqual([['Plain', 'none'], ['Bulleted', 'bullet']])
  })
})

describe('.pptx placeholder types', () => {
  it('takes a body placeholder\'s bullets from the master\'s body, not from its title that shares the layout\'s index', async () => {
    const box: SlideElement = { ...title, paragraphs: [paragraphOf('Bulleted')] }
    const parts = unzipSync(await writePptx({ ...presentation, slides: [{ id: newId(), background: '#ffffff', elements: [box], notes: '' }] }))
    parts['ppt/slides/slide1.xml'] = strToU8(strFromU8(parts['ppt/slides/slide1.xml'] ?? new Uint8Array())
      .replace('<p:nvPr></p:nvPr>', '<p:nvPr><p:ph idx="1"/></p:nvPr>')
      .replace(/<a:buNone\/>/g, ''))
    const placeholder = (id: number, type: string, idx: number, level: string): string => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${type} ${idx}"/><p:cNvSpPr/><p:nvPr><p:ph type="${type}" idx="${idx}"/></p:nvPr></p:nvSpPr>`
      + `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle>${level}</a:lstStyle><a:p/></p:txBody></p:sp>`
    const into = (name: string, shapes: string): void => {
      parts[name] = strToU8(strFromU8(parts[name] ?? new Uint8Array()).replace('</p:spTree>', `${shapes}</p:spTree>`))
    }
    into('ppt/slideLayouts/slideLayout1.xml', placeholder(2, 'body', 1, ''))
    // The master's title shares the layout body's index, and turns bullets off; its body has another index and bullets.
    into('ppt/slideMasters/slideMaster1.xml', placeholder(2, 'title', 1, '<a:lvl1pPr><a:buNone/></a:lvl1pPr>') + placeholder(3, 'body', 2, '<a:lvl1pPr><a:buChar char="•"/></a:lvl1pPr>'))

    const { presentation: reopened } = await readPptx(zipSync(parts))

    const text = reopened.slides[0]?.elements[0]
    expect(text?.kind === 'text' ? text.paragraphs[0]?.list : undefined).toBe('bullet')
  })
})

describe('.pptx table text colour', () => {
  it('takes a two-row table\'s text colour from its body, so turning the header row off keeps the header readable', async () => {
    const small = newTable(2, 2, { width: 960, height: 540, theme: 'light' })
    const deck: Presentation = { ...presentation, slides: [{ id: newId(), background: '#ffffff', elements: [small], notes: '' }] }

    const { presentation: reopened } = await readPptx(await writePptx(deck))

    const table = reopened.slides[0]?.elements[0]
    if (table?.kind !== 'table') throw new Error('The table did not come back')
    expect(table.color).toBe(small.color)
    expect(table.rows[0]?.cells[0]).toMatchObject({ color: '#ffffff', bold: true })
    expect(withHeaderRow(table, false, findTheme('light')).rows[0]?.cells[0]?.color ?? table.color).toBe(small.color)
  })
})

describe('.pptx line breaks', () => {
  it('keeps every line break inside a paragraph as a space, so no two words run together', async () => {
    const pptx = new PptxGenJS()
    pptx.addSlide().addText([
      { text: 'one' },
      { text: 'two', options: { softBreakBefore: true } },
      { text: 'three', options: { softBreakBefore: true, bold: true } },
      { text: 'four', options: { softBreakBefore: true } },
    ], { x: 0.5, y: 0.5, w: 6, h: 2 })
    const written = await pptx.write({ outputType: 'uint8array' })
    if (!(written instanceof Uint8Array)) throw new Error('pptxgenjs did not write bytes')

    const { presentation: imported, findings } = await readPptx(written)

    const box = imported.slides[0]?.elements[0]
    expect(box?.kind === 'shape' ? box.paragraphs[0]?.runs.map(run => [run.text, run.bold]) : undefined).toEqual([['one two ', false], ['three', true], [' four', false]])
    expect(findings.map(finding => finding.construct)).toEqual(['Line breaks inside paragraphs'])
  })

  it('keeps a tab and the line breaks together in one paragraph', async () => {
    const pptx = new PptxGenJS()
    pptx.addSlide().addText([
      { text: 'one\ttwo' },
      { text: 'three', options: { softBreakBefore: true, bold: true } },
      { text: 'four', options: { softBreakBefore: true } },
    ], { x: 0.5, y: 0.5, w: 6, h: 2 })
    const written = await pptx.write({ outputType: 'uint8array' })
    if (!(written instanceof Uint8Array)) throw new Error('pptxgenjs did not write bytes')

    const { presentation: imported } = await readPptx(written)

    const box = imported.slides[0]?.elements[0]
    expect(box?.kind === 'shape' ? box.paragraphs[0]?.runs.map(run => run.text).join('') : undefined).toBe('one\ttwo three four')
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
