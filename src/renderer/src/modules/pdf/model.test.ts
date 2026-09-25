import { describe, expect, it } from 'vitest'
import { changesSinceSave, markupBounds, rotateBox, turnPage, type PageItem, type Snapshot } from './model'

const page: PageItem = {
  key: 'p1',
  source: 0,
  index: 0,
  rotation: 0,
  markups: [
    { id: 'box', markup: { kind: 'rect', x: 10, y: 20, width: 100, height: 40, fill: '#ffff00', opacity: 0.35 } },
    { id: 'ink', markup: { kind: 'ink', points: [{ x: 5, y: 7 }, { x: 50, y: 70 }], width: 2, color: '#000000' } },
  ],
}
const portrait = { width: 600, height: 800 }
const landscape = { width: 800, height: 600 }

describe('turnPage', () => {
  it('keeps a box over the same content when the page turns clockwise', () => {
    const turned = turnPage(page, portrait, 1)
    expect(turned.rotation).toBe(90)
    // Content at the top-left of a portrait page ends up at the top-right of the landscape page.
    expect(turned.markups[0]?.markup).toMatchObject({ x: 800 - 20 - 40, y: 10, width: 40, height: 100 })
  })

  it('returns every markup to where it started after four turns either way', () => {
    let clockwise = page
    let counter = page
    for (let turn = 0; turn < 4; turn += 1) {
      clockwise = turnPage(clockwise, turn % 2 === 0 ? portrait : landscape, 1)
      counter = turnPage(counter, turn % 2 === 0 ? portrait : landscape, -1)
    }
    expect(clockwise).toEqual(page)
    expect(counter).toEqual(page)
  })

  it('keeps text centred over the same content, since text stays upright', () => {
    const text: PageItem = { ...page, markups: [{ id: 'text', markup: { kind: 'text', x: 10, y: 20, text: 'Hi', size: 10, color: '#000000' } }] }
    const before = markupBounds(text.markups[0]?.markup ?? page.markups[0].markup)
    const turned = turnPage(text, portrait, 1).markups[0]?.markup
    if (turned === undefined) throw new Error('missing markup')
    const after = markupBounds(turned)
    // The centre of the text follows the content: (x, y) on portrait → (800 - y, x) on landscape.
    expect(after.x + after.width / 2).toBeCloseTo(800 - (before.y + before.height / 2))
    expect(after.y + after.height / 2).toBeCloseTo(before.x + before.width / 2)
  })

  it('undoes a clockwise turn with a counter-clockwise one', () => {
    expect(turnPage(turnPage(page, portrait, 1), landscape, -1)).toEqual(page)
  })
})

describe('rotateBox', () => {
  it('matches turning the page one quarter at a time', () => {
    const box = { x: 10, y: 20, width: 100, height: 40 }
    expect(rotateBox(box, portrait, 90)).toEqual({ x: 740, y: 10, width: 40, height: 100 })
    expect(rotateBox(box, portrait, 180)).toEqual({ x: 490, y: 740, width: 100, height: 40 })
    expect(rotateBox(box, portrait, 0)).toEqual(box)
  })
})

describe('changesSinceSave', () => {
  const snapshot = (): Snapshot => ({ pages: [], formValues: {}, flattenForm: false })
  const a = snapshot()
  const b = snapshot()
  const c = snapshot()

  it('counts edits made after the save and edits undone past it', () => {
    expect(changesSinceSave([a], b, [], b)).toBe(0)
    expect(changesSinceSave([a, b], c, [], a)).toBe(2)
    // Saved at c, then undid twice: c is two redos away.
    expect(changesSinceSave([], a, [c, b], c)).toBe(2)
  })
})
