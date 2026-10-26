import type { FindMatch } from './slides-find'

const highlightName = 'find-match'

/** The characters `start` to `end` of the text inside `container`, counted across its text nodes. */
function textRange(container: Element, start: number, end: number): Range | undefined {
  const range = window.document.createRange()
  const nodes = window.document.createTreeWalker(container, NodeFilter.SHOW_TEXT)
  let offset = 0
  let started = false
  for (let node = nodes.nextNode(); node !== null; node = nodes.nextNode()) {
    const length = node.textContent?.length ?? 0
    if (!started && start < offset + length) {
      range.setStart(node, start - offset)
      started = true
    }
    if (started && end <= offset + length) {
      range.setEnd(node, end - offset)
      return range
    }
    offset += length
  }
  return undefined
}

/** Where a text box, shape or table-cell match is drawn inside `canvas`; undefined for notes, or when it is not drawn. */
export function matchRange(canvas: ParentNode, match: FindMatch): Range | undefined {
  if (match.where === 'notes') return undefined
  const element = canvas.querySelector(`[data-element-id="${match.elementId}"]`)
  const container = match.where === 'cell'
    ? element?.querySelector(`td[data-row="${match.row}"][data-column="${match.column}"]`)
    : element?.querySelectorAll('.slide-paragraph')[match.paragraph]
  return container === null || container === undefined ? undefined : textRange(container, match.start, match.end)
}

/**
 * Highlights a text box or table cell match on the slide, with focus staying where it is so Return in
 * the find bar steps on. Returns the function that takes the highlight away; a notes match has none.
 */
export function markMatch(match: FindMatch): () => void {
  const stage = window.document.querySelector('.slides-stage')
  const range = stage === null ? undefined : matchRange(stage, match)
  if (range === undefined) return () => undefined
  CSS.highlights.set(highlightName, new Highlight(range))
  return () => {
    CSS.highlights.delete(highlightName)
  }
}

/**
 * Selects a notes match in the notes field and scrolls it into view. The field takes focus: a field
 * without it does not show its selection.
 */
export function selectInNotes(match: FindMatch): void {
  const field = window.document.querySelector('.slides-notes')
  if (match.where !== 'notes' || !(field instanceof HTMLTextAreaElement)) return
  // Neither selecting nor focusing scrolls a textarea, so the match's height comes from the text before it.
  const text = field.value
  field.value = text.slice(0, match.end)
  const bottom = field.scrollHeight
  field.value = text
  // A short text is as high as the field, whatever it holds; then nothing needs scrolling.
  field.scrollTop = bottom <= field.clientHeight ? 0 : bottom - field.clientHeight / 2
  field.setSelectionRange(match.start, match.end)
  field.focus()
}
