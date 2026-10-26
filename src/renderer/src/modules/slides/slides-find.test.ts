import { describe, expect, it } from 'vitest'
import { newPresentation, newTable, paragraphOf, textRun, updateCell, updateSlide, type Presentation, type SlideElement, type TextBoxElement } from './model'
import { findMatches, occurrences, replaceMatches } from './slides-find'

function textBox(id: string, paragraphs: TextBoxElement['paragraphs']): TextBoxElement {
  return { kind: 'text', id, x: 0, y: 0, width: 100, height: 40, rotation: 0, paragraphs, verticalAlign: 'top', inset: { top: 0, right: 0, bottom: 0, left: 0 } }
}

/** Two slides: the first with a text box (the second paragraph is bold "Total" and plain " due"), a table and notes. */
function sample(): Presentation {
  const start = newPresentation()
  const table = updateCell(newTable(2, 2, start), 1, 1, { text: 'total: 40' })
  const elements: SlideElement[] = [
    textBox('box', [paragraphOf('Budget review'), { ...paragraphOf(''), runs: [textRun('Total', { bold: true }), textRun(' due')] }]),
    table,
  ]
  const first = updateSlide(start, start.slides[0]?.id ?? '', slide => ({ ...slide, elements, notes: 'Mention the total.' }))
  const second = { ...first.slides[0], id: 'second', elements: [textBox('other', [paragraphOf('TOTAL again')])], notes: '' }
  return { ...first, slides: [...first.slides, second] as Presentation['slides'] }
}

describe('findMatches', () => {
  it('finds text across slides, table cells and notes in reading order, ignoring case unless asked', () => {
    const presentation = sample()
    const found = findMatches(presentation, 'total', false)
    expect(found.map(match => match.where)).toEqual(['text', 'cell', 'notes', 'text'])
    expect(findMatches(presentation, 'Total', true)).toHaveLength(1)
    expect(findMatches(presentation, '', false)).toEqual([])
  })

  it('matches across runs of one paragraph but not across paragraphs', () => {
    const presentation = sample()
    expect(findMatches(presentation, 'Total due', true)).toHaveLength(1)
    expect(findMatches(presentation, 'review\nTotal', true)).toEqual([])
  })

  it('treats the query as plain text', () => {
    expect(occurrences('a.c abc', 'a.c', false)).toBe(1)
  })
})

describe('replaceMatches', () => {
  it('replaces every match in one new presentation, keeping the run styles', () => {
    const presentation = sample()
    const replaced = replaceMatches(presentation, findMatches(presentation, 'total', false), '$&sum')
    const [first, second] = replaced.slides
    const box = first?.elements[0]
    if (box?.kind !== 'text') throw new Error('No text box')
    expect(box.paragraphs[1]?.runs).toMatchObject([{ text: '$&sum', bold: true }, { text: ' due', bold: false }])
    const table = first?.elements[1]
    expect(table?.kind === 'table' ? table.rows[1]?.cells[1]?.text : undefined).toBe('$&sum: 40')
    expect(first?.notes).toBe('Mention the $&sum.')
    expect(second?.elements[0]).toMatchObject({ paragraphs: [{ runs: [{ text: '$&sum again' }] }] })
    expect(findMatches(replaced, 'total', false)).toEqual([])
  })

  it('replaces a match that spans runs with the first run\'s style, and one match at a time', () => {
    const presentation = sample()
    const [match] = findMatches(presentation, 'Total due', true)
    if (match === undefined) throw new Error('No match')
    const replaced = replaceMatches(presentation, [match], 'Paid')
    const box = replaced.slides[0]?.elements[0]
    expect(box?.kind === 'text' ? box.paragraphs[1]?.runs : undefined).toMatchObject([{ text: 'Paid', bold: true }])
    expect(findMatches(replaced, 'total', false)).toHaveLength(3)
  })

  it('keeps an empty run in a paragraph that loses all its text', () => {
    const presentation = sample()
    const [match] = findMatches(presentation, 'Budget review', true)
    if (match === undefined) throw new Error('No match')
    const box = replaceMatches(presentation, [match], '').slides[0]?.elements[0]
    expect(box?.kind === 'text' ? box.paragraphs[0]?.runs : undefined).toMatchObject([{ text: '' }])
  })
})
