// @vitest-environment happy-dom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it } from 'vitest'
import { matchRange } from './find-reveal'
import { newPresentation, newTable, paragraphOf, textRun, updateCell, updateSlide, type TextBoxElement } from './model'
import { SlideView } from './SlideView'
import { findMatches } from './slides-find'

Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })

function textBox(paragraphs: TextBoxElement['paragraphs']): TextBoxElement {
  return { kind: 'text', id: 'box', x: 0, y: 0, width: 100, height: 40, rotation: 0, paragraphs, verticalAlign: 'top', inset: { top: 0, right: 0, bottom: 0, left: 0 } }
}

it('finds where each occurrence is drawn, in a text box across runs and in a table cell', async () => {
  const start = newPresentation()
  const table = updateCell(newTable(1, 2, start), 0, 1, { text: 'ab ab' })
  const slide = updateSlide(start, start.slides[0]?.id ?? '', current => ({
    ...current,
    elements: [textBox([paragraphOf('first'), { ...paragraphOf(''), runs: [textRun('go g', { bold: true }), textRun('o go')] }]), table],
  })).slides[0]
  if (slide === undefined) throw new Error('No slide')
  const host = document.createElement('div')
  document.body.append(host)
  const root = createRoot(host)
  await act(async () => root.render(<SlideView slide={slide} width={960} height={540} scale={1} />))

  const presentation = { ...start, slides: [slide] }
  const drawn = (query: string): Array<readonly [string, number, string]> => findMatches(presentation, query, true).map(match => {
    const range = matchRange(host, match)
    return [range?.toString() ?? '', range?.startOffset ?? -1, range?.startContainer.parentElement?.tagName ?? ''] as const
  })
  // The second "go" starts in the first run and ends in the second; the third starts after it in the second run.
  expect(drawn('go').map(([text]) => text)).toEqual(['go', 'go', 'go'])
  expect(drawn('go').map(([, offset]) => offset)).toEqual([0, 3, 2])
  expect(drawn('ab')).toEqual([['ab', 0, 'TD'], ['ab', 3, 'TD']])
  await act(async () => root.unmount())
  host.remove()
})
