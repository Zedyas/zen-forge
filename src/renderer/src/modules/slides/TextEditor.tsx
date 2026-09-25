import { memo, useLayoutEffect, useRef, useState } from 'react'
import {
  bulletIndent,
  defaultFont,
  holdsText,
  parseColor,
  paragraphOf,
  shapeTextStyle,
  type HorizontalAlign,
  type Paragraph,
  type RunStyle,
  type ShapeElement,
  type TextBoxElement,
  type TextRun,
} from './model'
import { changeElement, removeElement } from './slides-actions'
import { readySlides, updateSlides } from './slides-store'
import { ElementView, TextParagraphs } from './SlideView'
import { registerEditor, type InlineFormat } from './text-editing'

/*
 * Inline text editing: while a box is edited, its text is a contenteditable element laid out exactly
 * like the static slide. The browser's editing commands apply bold, italic, underline, colour, size
 * and alignment; closing the editor reads the markup back into paragraphs and runs.
 */

const blockTags: ReadonlySet<string> = new Set(['P', 'DIV', 'LI', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE', 'UL', 'OL'])

function isBlock(node: Node): node is HTMLElement {
  return node instanceof HTMLElement && blockTags.has(node.tagName)
}

/** A run's style as the browser computes it, which covers every way editing commands mark it up. */
function styleAt(element: Element, root: HTMLElement): RunStyle {
  const style = getComputedStyle(element)
  let underline = false
  for (let node: Element | null = element; node !== null && root.contains(node); node = node.parentElement) {
    if (getComputedStyle(node).textDecorationLine.includes('underline')) underline = true
  }
  return {
    bold: Number(style.fontWeight) >= 600,
    italic: style.fontStyle !== 'normal',
    underline,
    color: parseColor(style.color) ?? '#000000',
    // The editor is laid out at one CSS pixel per point.
    size: Math.round(Number.parseFloat(style.fontSize) * 10) / 10,
    font: style.fontFamily.split(',')[0]?.replace(/["']/g, '').trim() || defaultFont,
  }
}

function sameStyle(a: TextRun, b: TextRun): boolean {
  return a.bold === b.bold && a.italic === b.italic && a.underline === b.underline && a.color === b.color && a.size === b.size && a.font === b.font
}

function merged(runs: readonly TextRun[]): TextRun[] {
  return runs.reduce<TextRun[]>((all, run) => {
    const last = all.at(-1)
    if (last !== undefined && sameStyle(last, run)) all[all.length - 1] = { ...last, text: last.text + run.text }
    else all.push(run)
    return all
  }, [])
}

function hasContentAfter(node: Node): boolean {
  for (let next = node.nextSibling; next !== null; next = next.nextSibling) {
    if (next instanceof HTMLBRElement || (next.textContent ?? '') !== '') return true
  }
  return false
}

/** Chromium ends an empty line with a `<br>`; only a `<br>` followed by more text breaks a paragraph. */
function isTrailingBreak(br: HTMLBRElement, root: HTMLElement): boolean {
  let node: Node = br
  while (node.parentNode !== null && node.parentNode !== root && !isBlock(node.parentNode)) {
    if (hasContentAfter(node)) return false
    node = node.parentNode
  }
  return !hasContentAfter(node)
}

function alignOf(value: string): HorizontalAlign {
  if (value === 'center' || value === 'justify') return value
  return value === 'right' || value === 'end' ? 'right' : 'left'
}

type ParagraphProps = Omit<Paragraph, 'runs'>

function blockProps(block: HTMLElement, previous: ParagraphProps): ParagraphProps {
  const level = Number(block.dataset['level'])
  return {
    align: alignOf(getComputedStyle(block).textAlign),
    bullet: block.tagName === 'LI' || block.classList.contains('is-bullet'),
    level: Number.isInteger(level) ? Math.min(8, Math.max(0, level)) : previous.level,
  }
}

/** Reads the editor's markup back into paragraphs; `first` supplies the properties of text outside any block. */
function readEditor(root: HTMLElement, first: Paragraph): Paragraph[] {
  const paragraphs: Paragraph[] = []
  let runs: TextRun[] = []
  let props: ParagraphProps = { align: first.align, bullet: first.bullet, level: first.level }
  let emptyStyle: RunStyle = styleAt(root, root)
  const flush = (): void => {
    paragraphs.push({ ...props, runs: runs.length > 0 ? merged(runs) : [{ text: '', ...emptyStyle }] })
    runs = []
  }
  const walk = (node: Node): void => {
    if (node instanceof Text) {
      if (node.data !== '' && node.parentElement !== null) runs.push({ text: node.data.replace(/\u00a0/g, ' '), ...styleAt(node.parentElement, root) })
      return
    }
    if (!(node instanceof HTMLElement)) return
    if (node instanceof HTMLBRElement) {
      if (node.parentElement !== null) emptyStyle = styleAt(node.parentElement, root)
      if (!isTrailingBreak(node, root)) flush()
      return
    }
    const block = isBlock(node)
    if (block) {
      if (runs.length > 0) flush()
      props = blockProps(node, props)
      emptyStyle = styleAt(node, root)
    }
    node.childNodes.forEach(walk)
    if (block && (runs.length > 0 || node.querySelector(Array.from(blockTags).join(',')) === null)) flush()
  }
  root.childNodes.forEach(walk)
  if (runs.length > 0 || paragraphs.length === 0) flush()
  return paragraphs
}

/** Gives each line its first run's font, as the static drawing does, so line heights and bullets follow size changes. */
function syncLineStyles(root: HTMLElement): void {
  root.querySelectorAll<HTMLElement>('.slide-paragraph, p, div, li').forEach(block => {
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT)
    for (let node = walker.nextNode(); node !== null; node = walker.nextNode()) {
      if ((node.textContent ?? '') === '' || node.parentElement === null) continue
      const style = getComputedStyle(node.parentElement)
      block.style.fontSize = style.fontSize
      block.style.fontFamily = style.fontFamily
      block.style.color = style.color
      return
    }
  })
}

function commandFor(align: HorizontalAlign): string {
  return { left: 'justifyLeft', center: 'justifyCenter', right: 'justifyRight', justify: 'justifyFull' }[align]
}

/**
 * Formats the selected text. The live selection wins while it is still in the editor (a menu
 * shortcut); after a toolbar control took focus, the selection saved before that is put back.
 */
function applyFormat(root: HTMLElement, saved: Range | undefined, change: InlineFormat): void {
  const selection = window.getSelection()
  const inside = selection !== null && selection.rangeCount > 0 && root.contains(selection.anchorNode)
  root.focus()
  if (!inside && selection !== null && saved !== undefined) {
    selection.removeAllRanges()
    selection.addRange(saved)
  }
  document.execCommand('styleWithCSS', false, 'true')
  switch (change.kind) {
    case 'toggle':
      document.execCommand(change.style)
      break
    case 'color':
      document.execCommand('foreColor', false, change.value)
      break
    case 'align':
      document.execCommand(commandFor(change.value))
      break
    case 'size': {
      // fontSize only takes the HTML sizes 1 to 7: mark with 7, then replace the marks with the exact size.
      // With nothing selected the whole box changes, since a caret alone cannot hold a size.
      if (selection !== null && selection.isCollapsed) selection.selectAllChildren(root)
      document.execCommand('styleWithCSS', false, 'false')
      document.execCommand('fontSize', false, '7')
      root.querySelectorAll('font[size="7"]').forEach(font => {
        const span = document.createElement('span')
        span.style.fontSize = `${change.value}px`
        span.append(...Array.from(font.childNodes))
        font.replaceWith(span)
      })
      break
    }
  }
  syncLineStyles(root)
}

/** Moves the paragraphs under the caret one bullet level in or out. */
function indent(root: HTMLElement, step: 1 | -1): void {
  const anchor = window.getSelection()?.anchorNode ?? null
  let block: Node | null = anchor
  while (block !== null && block.parentNode !== root) block = block.parentNode
  if (!(block instanceof HTMLElement)) return
  const level = Math.min(8, Math.max(0, (Number(block.dataset['level']) || 0) + step))
  block.dataset['level'] = String(level)
  if (block.classList.contains('is-bullet')) block.style.paddingLeft = `${bulletIndent * (level + 1)}px`
}

function placeCaret(root: HTMLElement, point: { readonly x: number; readonly y: number } | undefined): void {
  const selection = window.getSelection()
  if (selection === null) return
  const range = point === undefined ? null : document.caretRangeFromPoint(point.x, point.y)
  if (range !== null && root.contains(range.startContainer)) {
    selection.removeAllRanges()
    selection.addRange(range)
    return
  }
  selection.selectAllChildren(root)
  selection.collapseToEnd()
}

/** Writes typed text into the element, unless nothing changed. An emptied new text box is removed, as in Keynote. */
function writeText(documentId: string, elementId: string, paragraphs: readonly Paragraph[]): void {
  const element = readySlides(documentId)?.present.slides.flatMap(slide => slide.elements).find(candidate => candidate.id === elementId)
  if (element === undefined || !holdsText(element)) return
  const empty = paragraphs.every(paragraph => paragraph.runs.every(run => run.text === ''))
  if (empty && element.kind === 'text' && element.prompt === undefined && element.fill === undefined && element.border === undefined) {
    removeElement(documentId, elementId)
    return
  }
  const next = element.kind === 'shape' && empty ? [] : paragraphs
  if (JSON.stringify(next) === JSON.stringify(element.paragraphs)) return
  changeElement(documentId, elementId, current => holdsText(current) ? { ...current, paragraphs: next } : current)
}

/** The first paragraph's markup, never re-rendered: after mounting, the browser owns the editor's content. */
const InitialText = memo(function InitialText({ paragraphs }: { readonly paragraphs: readonly Paragraph[] }) {
  return <TextParagraphs paragraphs={paragraphs} />
}, () => true)

interface TextEditorProps {
  readonly documentId: string
  readonly element: TextBoxElement | ShapeElement
  /** Where the double-click that opened the editor landed, in window coordinates, for the caret. */
  readonly caret?: { readonly x: number; readonly y: number }
}

export function TextEditor({ documentId, element, caret }: TextEditorProps) {
  const rootRef = useRef<HTMLDivElement>(null)
  const saved = useRef<Range | undefined>(undefined)
  const [initial] = useState(() => element.paragraphs.length > 0 ? element.paragraphs : [paragraphOf('', shapeTextStyle(element.fill), { align: 'center' })])
  const close = useRef<() => void>(() => undefined)

  useLayoutEffect(() => {
    const root = rootRef.current
    const first = initial[0]
    if (root === null || first === undefined) return
    let written = false
    const write = (): void => {
      if (written) return
      written = true
      writeText(documentId, element.id, readEditor(root, first))
    }
    close.current = () => {
      write()
      updateSlides(documentId, document => document.editingId === element.id ? { editingId: undefined } : {})
    }
    document.execCommand('defaultParagraphSeparator', false, 'p')
    root.focus()
    placeCaret(root, caret)
    const remember = (): void => {
      const selection = window.getSelection()
      if (selection !== null && selection.rangeCount > 0 && root.contains(selection.anchorNode)) saved.current = selection.getRangeAt(0).cloneRange()
    }
    document.addEventListener('selectionchange', remember)
    const unregister = registerEditor({
      documentId,
      elementId: element.id,
      format: change => applyFormat(root, saved.current, change),
      commit: () => close.current(),
    })
    // Switching tabs or slides unmounts the editor without a blur; what was typed is kept.
    return () => {
      document.removeEventListener('selectionchange', remember)
      unregister()
      write()
    }
  }, [])

  return (
    <ElementView
      element={element}
      text={(
        <div
          ref={rootRef}
          className="slide-text slide-text-editor"
          contentEditable
          suppressContentEditableWarning
          spellCheck
          role="textbox"
          aria-multiline="true"
          aria-label="Slide text"
          onPointerDown={event => event.stopPropagation()}
          onBlur={event => {
            // Toolbar menus and colour pickers take focus while they are open; the editor stays.
            const next = event.relatedTarget
            if (next instanceof Element && next.closest('.toolbar, .popover, .menu-popup, .title-essentials') !== null) return
            close.current()
          }}
          onPaste={event => {
            // Pasted text takes the formatting where it lands, as PowerPoint's "Keep Text Only".
            event.preventDefault()
            document.execCommand('insertText', false, event.clipboardData.getData('text/plain'))
          }}
          onKeyDown={event => {
            if (event.key === 'Escape') {
              event.preventDefault()
              close.current()
            } else if (event.key === 'Tab') {
              event.preventDefault()
              if (rootRef.current !== null) indent(rootRef.current, event.shiftKey ? -1 : 1)
            }
          }}
        >
          <InitialText paragraphs={initial} />
        </div>
      )}
    />
  )
}
