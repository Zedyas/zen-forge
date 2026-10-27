import { Extension, type Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'

export interface FindMatch {
  readonly from: number
  readonly to: number
}

export interface FindState {
  readonly query: string
  readonly caseSensitive: boolean
  readonly matches: readonly FindMatch[]
  /** Index of the current match in `matches`. */
  readonly current: number
  /** Text that replacements of this search inserted; it is never matched, so Replace cannot re-enter it. */
  readonly inserted: readonly FindMatch[]
  readonly decorations: DecorationSet
}

/** `inserted` is new text a transaction put in, in the positions of the document after it. */
type FindUpdate = Partial<Pick<FindState, 'query' | 'caseSensitive' | 'current'>> & { readonly inserted?: readonly FindMatch[] }

const findKey = new PluginKey<FindState>('find')

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/** Matches within each paragraph; an image or line break ends a match, as in Word. */
export function findMatches(doc: ProseMirrorNode, query: string, caseSensitive: boolean): FindMatch[] {
  if (query === '') return []
  const pattern = new RegExp(escapeRegExp(query), caseSensitive ? 'g' : 'gi')
  const matches: FindMatch[] = []
  doc.descendants((node, position) => {
    if (!node.isTextblock) return true
    // Every inline node that is not text is one position wide, so one placeholder character keeps offsets aligned.
    let text = ''
    node.forEach(child => {
      text += child.isText ? child.text ?? '' : '￼'
    })
    for (const match of text.matchAll(pattern)) {
      const from = position + 1 + match.index
      matches.push({ from, to: from + match[0].length })
    }
    return false
  })
  return matches
}

function decorate(doc: ProseMirrorNode, matches: readonly FindMatch[], current: number): DecorationSet {
  return DecorationSet.create(doc, matches.map((match, index) =>
    Decoration.inline(match.from, match.to, { class: index === current ? 'sumi-find-match is-current' : 'sumi-find-match' })))
}

function build(doc: ProseMirrorNode, query: string, caseSensitive: boolean, current: number, inserted: readonly FindMatch[]): FindState {
  const matches = findMatches(doc, query, caseSensitive).filter(match => !inserted.some(range => match.from < range.to && match.to > range.from))
  const index = matches.length === 0 ? 0 : Math.min(Math.max(current, 0), matches.length - 1)
  return { query, caseSensitive, matches, current: index, inserted, decorations: decorate(doc, matches, index) }
}

/** Highlights find matches; the find bar drives it through the functions below. */
export const FindHighlight = Extension.create({
  name: 'findHighlight',
  addProseMirrorPlugins() {
    return [new Plugin<FindState>({
      key: findKey,
      state: {
        init: (_config, state) => build(state.doc, '', false, 0, []),
        apply: (transaction, previous, _oldState, state) => {
          const update: FindUpdate | undefined = transaction.getMeta(findKey)
          if (update === undefined && !transaction.docChanged) return previous
          // A new search starts over; otherwise the inserted text follows the edits around it.
          const inserted = update?.query !== undefined || update?.caseSensitive !== undefined ? [] : [
            ...previous.inserted
              .map(range => ({ from: transaction.mapping.map(range.from, 1), to: transaction.mapping.map(range.to, -1) }))
              .filter(range => range.to > range.from),
            ...update?.inserted ?? [],
          ]
          return build(state.doc, update?.query ?? previous.query, update?.caseSensitive ?? previous.caseSensitive, update?.current ?? previous.current, inserted)
        },
      },
      props: {
        decorations: state => findKey.getState(state)?.decorations,
      },
    })]
  },
})

export function findState(state: EditorState): FindState | undefined {
  return findKey.getState(state)
}

function update(editor: Editor, change: FindUpdate): void {
  editor.view.dispatch(editor.state.tr.setMeta(findKey, change))
}

/** Scrolls the current match into the middle of the desk. */
function revealCurrent(editor: Editor): void {
  editor.view.dom.querySelector('.sumi-find-match.is-current')?.scrollIntoView({ block: 'center' })
}

export function setFindQuery(editor: Editor, query: string, caseSensitive: boolean): void {
  update(editor, { query, caseSensitive, current: 0 })
  revealCurrent(editor)
}

export function moveFind(editor: Editor, step: 1 | -1): void {
  const state = findState(editor.state)
  if (state === undefined || state.matches.length === 0) return
  update(editor, { current: (state.current + step + state.matches.length) % state.matches.length })
  revealCurrent(editor)
}

/**
 * Replaces the current match; the next match after the new text becomes current. The new text is never
 * matched by this search, so a replacement that holds the search text is not replaced again, and
 * Replace does nothing once no other match is left. Keeps the formatting of the replaced text.
 */
export function replaceCurrent(editor: Editor, replacement: string): void {
  const state = findState(editor.state)
  const match = state?.matches[state.current]
  if (match === undefined) return
  const end = match.from + replacement.length
  const transaction = replacement === '' ? editor.state.tr.delete(match.from, match.to) : editor.state.tr.insertText(replacement, match.from, match.to)
  editor.view.dispatch(transaction.setMeta(findKey, { inserted: replacement === '' ? [] : [{ from: match.from, to: end }] }))
  const next = findState(editor.state)?.matches.findIndex(candidate => candidate.from >= end) ?? -1
  update(editor, { current: Math.max(next, 0) })
  revealCurrent(editor)
}

/** Replaces every match in one step, so one Undo restores them all. Resolves the count. */
export function replaceAll(editor: Editor, replacement: string): number {
  const matches = findState(editor.state)?.matches ?? []
  const transaction = editor.state.tr
  // Last to first, so earlier positions stay valid.
  for (const match of [...matches].reverse()) {
    if (replacement === '') transaction.delete(match.from, match.to)
    else transaction.insertText(replacement, match.from, match.to)
  }
  // Where each replacement ends up, once the ones before it have changed the length.
  let shift = 0
  const inserted = replacement === '' ? [] : matches.map(match => {
    const from = match.from + shift
    shift += replacement.length - (match.to - match.from)
    return { from, to: from + replacement.length }
  })
  if (matches.length > 0) editor.view.dispatch(transaction.setMeta(findKey, { inserted }))
  return matches.length
}

export function clearFind(editor: Editor): void {
  if (!editor.isDestroyed) update(editor, { query: '' })
}
