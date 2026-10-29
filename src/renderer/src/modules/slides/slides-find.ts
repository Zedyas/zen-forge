import { holdsText, mergeRuns, updateCell, updateElement, updateSlide, type Paragraph, type Presentation, type TextRun } from './model'

/**
 * Where a match is: in one paragraph of a text box or shape, in a table cell, or in a slide's
 * speaker notes. `start` and `end` count characters of that one text.
 */
export type FindMatch =
  | { readonly where: 'text'; readonly slideId: string; readonly elementId: string; readonly paragraph: number; readonly start: number; readonly end: number }
  | { readonly where: 'cell'; readonly slideId: string; readonly elementId: string; readonly row: number; readonly column: number; readonly start: number; readonly end: number }
  | { readonly where: 'notes'; readonly slideId: string; readonly start: number; readonly end: number }

interface Range {
  readonly start: number
  readonly end: number
}

function ranges(text: string, query: string, caseSensitive: boolean): Range[] {
  if (query === '') return []
  const pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), caseSensitive ? 'g' : 'gi')
  return Array.from(text.matchAll(pattern), match => ({ start: match.index, end: match.index + match[0].length }))
}

/** How many times `query` occurs in `text`. */
export function occurrences(text: string, query: string, caseSensitive: boolean): number {
  return ranges(text, query, caseSensitive).length
}

function paragraphText(paragraph: Paragraph): string {
  return paragraph.runs.map(run => run.text).join('')
}

/**
 * Every match in the presentation, in reading order: each slide's text boxes, shapes and tables back
 * to front, then its notes. A match never spans paragraphs or cells.
 */
export function findMatches(presentation: Presentation, query: string, caseSensitive: boolean): FindMatch[] {
  return presentation.slides.flatMap(slide => [
    ...slide.elements.flatMap((element): FindMatch[] => {
      if (holdsText(element)) {
        return element.paragraphs.flatMap((paragraph, index) =>
          ranges(paragraphText(paragraph), query, caseSensitive).map(range => ({ where: 'text' as const, slideId: slide.id, elementId: element.id, paragraph: index, ...range })))
      }
      if (element.kind !== 'table') return []
      return element.rows.flatMap((row, rowIndex) => row.cells.flatMap((cell, column) =>
        ranges(cell.text, query, caseSensitive).map(range => ({ where: 'cell' as const, slideId: slide.id, elementId: element.id, row: rowIndex, column, ...range }))))
    }),
    ...ranges(slide.notes, query, caseSensitive).map(range => ({ where: 'notes' as const, slideId: slide.id, ...range })),
  ])
}

function spliceText(text: string, range: Range, replacement: string): string {
  return text.slice(0, range.start) + replacement + text.slice(range.end)
}

/** Puts `replacement` where the match was, in the style of the run the match starts in. */
function spliceRuns(runs: readonly TextRun[], range: Range, replacement: string): TextRun[] {
  let offset = 0
  const spliced = runs.map(run => {
    const from = offset
    offset += run.text.length
    const within = (index: number): number => Math.min(Math.max(index - from, 0), run.text.length)
    const inserted = from <= range.start && range.start < offset ? replacement : ''
    return { ...run, text: run.text.slice(0, within(range.start)) + inserted + run.text.slice(within(range.end)) }
  })
  const kept = spliced.filter(run => run.text !== '')
  // A paragraph keeps one empty run, which carries its size and font.
  return mergeRuns(kept.length > 0 ? kept : spliced.slice(0, 1))
}

function replaceOne(presentation: Presentation, match: FindMatch, replacement: string): Presentation {
  return updateSlide(presentation, match.slideId, slide => {
    if (match.where === 'notes') return { ...slide, notes: spliceText(slide.notes, match, replacement) }
    return updateElement(slide, match.elementId, element => {
      if (match.where === 'text' && holdsText(element)) {
        return {
          ...element,
          paragraphs: element.paragraphs.map((paragraph, index) =>
            index === match.paragraph ? { ...paragraph, runs: spliceRuns(paragraph.runs, match, replacement) } : paragraph),
        }
      }
      if (match.where === 'cell' && element.kind === 'table') {
        const text = element.rows[match.row]?.cells[match.column]?.text
        return text === undefined ? element : updateCell(element, match.row, match.column, { text: spliceText(text, match, replacement) })
      }
      return element
    })
  })
}

/** Replaces matches found in `presentation`, last first so earlier offsets in the same text stay valid. */
export function replaceMatches(presentation: Presentation, matches: readonly FindMatch[], replacement: string): Presentation {
  return matches.reduceRight((current, match) => replaceOne(current, match, replacement), presentation)
}
