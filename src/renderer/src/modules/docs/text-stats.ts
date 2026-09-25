import type { Node as ProseMirrorNode } from '@tiptap/pm/model'

export interface TextCounts {
  readonly words: number
  readonly characters: number
}

export interface OutlineHeading {
  readonly level: number
  readonly text: string
  /** Position just inside the heading, where the caret goes to jump there. */
  readonly position: number
}

const counts = new WeakMap<ProseMirrorNode, TextCounts>()
const wordSegmenter = new Intl.Segmenter(undefined, { granularity: 'word' })

/** Words and characters (spaces included), as Word and Google Docs count them. Cached per document state. */
export function countText(doc: ProseMirrorNode): TextCounts {
  const cached = counts.get(doc)
  if (cached !== undefined) return cached
  const text = doc.textBetween(0, doc.content.size, '\n', ' ')
  let words = 0
  for (const segment of wordSegmenter.segment(text)) {
    if (segment.isWordLike === true) words += 1
  }
  const result = { words, characters: text.replaceAll('\n', '').length }
  counts.set(doc, result)
  return result
}

/** The document's title and headings in order, for the inspector's outline. Titles are level 0. */
export function outline(doc: ProseMirrorNode): OutlineHeading[] {
  const headings: OutlineHeading[] = []
  doc.descendants((node, position) => {
    if (node.type.name === 'title') headings.push({ level: 0, text: node.textContent, position: position + 1 })
    if (node.type.name === 'heading') {
      const level: unknown = node.attrs['level']
      headings.push({ level: typeof level === 'number' ? level : 1, text: node.textContent, position: position + 1 })
    }
    return !node.isTextblock
  })
  return headings
}
