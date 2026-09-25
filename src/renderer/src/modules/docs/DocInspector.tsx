import { useEffect, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { useEditorState } from '@tiptap/react'
import { Inspector, InspectorRow, InspectorSection } from '../../ui/Inspector'
import { Tip } from '../../ui/Tip'
import { currentBlockStyle, setPage, setParagraphFormat, type ParagraphFormat } from './doc-commands'
import { countText, outline } from './text-stats'
import { documentPage, isLandscape, lengthUnit, pageFor, paperOf, papers, type PageSetup } from './theme'

const pointsPerUnit = lengthUnit === 'in' ? 72 : 72 / 2.54

function formatNumber(value: number): string {
  return String(Math.round(value * 100) / 100)
}

/** A number typed in a text field, applied on Enter or when the field loses focus. */
function NumberField({ label, value, unit, min, max, onCommit }: {
  readonly label: string
  readonly value: number
  readonly unit: string
  readonly min: number
  readonly max: number
  onCommit(value: number): void
}) {
  const [draft, setDraft] = useState(formatNumber(value))
  useEffect(() => setDraft(formatNumber(value)), [value])
  const commit = (): void => {
    const parsed = Number(draft.replace(',', '.'))
    if (Number.isFinite(parsed) && parsed >= min && parsed <= max) onCommit(parsed)
    else setDraft(formatNumber(value))
  }
  return (
    <InspectorRow label={label}>
      <span className="sumi-number">
        <input
          className="text-field"
          inputMode="decimal"
          value={draft}
          aria-label={`${label} (${unit})`}
          onChange={event => setDraft(event.currentTarget.value)}
          onBlur={commit}
          onKeyDown={event => {
            if (event.key === 'Enter') commit()
            if (event.key === 'Escape') setDraft(formatNumber(value))
          }}
        />
        <small>{unit}</small>
      </span>
    </InspectorRow>
  )
}

interface ParagraphState {
  readonly editable: boolean
  readonly inList: boolean
  readonly spaceBefore: number
  readonly spaceAfter: number
  readonly lineHeight: number
  readonly indentLeft: number
  readonly indentFirstLine: number
  readonly base: ReturnType<typeof currentBlockStyle>
}

function paragraphState(editor: Editor): ParagraphState {
  const block = editor.state.selection.$from.parent
  const base = currentBlockStyle(editor)
  const value = (name: string, fallback: number): number => {
    const attr: unknown = block.attrs[name]
    return typeof attr === 'number' ? attr : fallback
  }
  return {
    editable: ['paragraph', 'heading', 'title'].includes(block.type.name),
    inList: editor.isActive('listItem') || editor.isActive('taskItem'),
    spaceBefore: value('spaceBefore', base.spaceBefore),
    spaceAfter: value('spaceAfter', base.spaceAfter),
    lineHeight: value('lineHeight', base.lineHeight),
    indentLeft: value('indentLeft', base.indentLeft),
    indentFirstLine: value('indentFirstLine', 0),
    base,
  }
}

function ParagraphSection({ editor }: { readonly editor: Editor }) {
  const state = useEditorState({ editor, selector: ({ editor: current }) => paragraphState(current) })
  if (!state.editable) return null
  // A value equal to the style's is no direct formatting, so the paragraph follows its style.
  const set = (name: keyof ParagraphFormat, value: number, styleValue: number): void => setParagraphFormat(editor, { [name]: value === styleValue ? null : value })
  return (
    <InspectorSection title="Paragraph">
      <NumberField label="Space before" value={state.spaceBefore} unit="pt" min={0} max={200} onCommit={value => set('spaceBefore', value, state.base.spaceBefore)} />
      <NumberField label="Space after" value={state.spaceAfter} unit="pt" min={0} max={200} onCommit={value => set('spaceAfter', value, state.base.spaceAfter)} />
      <NumberField label="Line spacing" value={state.lineHeight} unit="lines" min={0.5} max={5} onCommit={value => set('lineHeight', value, state.base.lineHeight)} />
      {state.inList
        ? <p className="inspector-note">List items take their indent from the list level. Use Tab and Shift-Tab to change it.</p>
        : (
          <>
            <NumberField label="Left indent" value={state.indentLeft / pointsPerUnit} unit={lengthUnit} min={0} max={20} onCommit={value => set('indentLeft', Math.round(value * pointsPerUnit * 10) / 10, state.base.indentLeft)} />
            <NumberField label="First line" value={state.indentFirstLine / pointsPerUnit} unit={lengthUnit} min={-20} max={20} onCommit={value => set('indentFirstLine', Math.round(value * pointsPerUnit * 10) / 10, 0)} />
            <p className="inspector-note">A negative first line makes a hanging indent.</p>
          </>
        )}
    </InspectorSection>
  )
}

const cellShades: ReadonlyArray<{ readonly label: string; readonly value: string | null }> = [
  { label: 'No shading', value: null },
  { label: 'Grey', value: '#f2f2f2' },
  { label: 'Yellow', value: '#fff2cc' },
  { label: 'Green', value: '#e2efd9' },
  { label: 'Blue', value: '#deeaf6' },
  { label: 'Red', value: '#fbe4d5' },
  { label: 'Purple', value: '#e5dfec' },
]

function tableAt(editor: Editor): ProseMirrorNode | undefined {
  const { $from } = editor.state.selection
  for (let depth = $from.depth; depth > 0; depth -= 1) {
    if ($from.node(depth).type.name === 'table') return $from.node(depth)
  }
  return undefined
}

function TableSection({ editor }: { readonly editor: Editor }) {
  const state = useEditorState({
    editor,
    selector: ({ editor: current }) => {
      const table = tableAt(current)
      return {
        inTable: table !== undefined,
        headerRow: table?.firstChild?.firstChild?.type.name === 'tableHeader',
        canMerge: current.can().mergeCells(),
        canSplit: current.can().splitCell(),
      }
    },
  })
  if (!state.inTable) return null
  const chain = () => editor.chain().focus()
  return (
    <InspectorSection title="Table">
      <label className="sumi-check-row">
        <input type="checkbox" checked={state.headerRow} onChange={() => chain().toggleHeaderRow().run()} />
        Header row
      </label>
      <div className="sumi-inspector-actions">
        <button type="button" className="button" disabled={!state.canMerge} onClick={() => chain().mergeCells().run()}>Merge cells</button>
        <button type="button" className="button" disabled={!state.canSplit} onClick={() => chain().splitCell().run()}>Split cell</button>
      </div>
      <InspectorRow label="Cell shading"><span /></InspectorRow>
      <div className="swatch-grid sumi-shades">
        {cellShades.map(shade => (
          <Tip key={shade.label} label={shade.label}>
            <button
              type="button"
              className={`swatch${shade.value === null ? ' is-none' : ''}`}
              style={shade.value === null ? undefined : { background: shade.value }}
              aria-label={shade.label}
              onClick={() => chain().setCellAttribute('backgroundColor', shade.value).run()}
            />
          </Tip>
        ))}
      </div>
      <p className="inspector-note">Drag a column's border to change its width.</p>
    </InspectorSection>
  )
}

function PageSection({ editor }: { readonly editor: Editor }) {
  const page = useEditorState({ editor, selector: ({ editor: current }) => documentPage(current.state.doc.attrs['page']) })
  const paper = paperOf(page)
  const landscape = isLandscape(page)
  const margin = (name: keyof Pick<PageSetup, 'marginTop' | 'marginBottom' | 'marginLeft' | 'marginRight'>, label: string) => (
    <NumberField
      label={label}
      value={page[name] / 1440 * (lengthUnit === 'in' ? 1 : 2.54)}
      unit={lengthUnit}
      min={0}
      max={lengthUnit === 'in' ? 4 : 10}
      onCommit={value => setPage(editor, { [name]: Math.round(value / (lengthUnit === 'in' ? 1 : 2.54) * 1440) })}
    />
  )
  return (
    <InspectorSection title="Page">
      <InspectorRow label="Paper">
        <span className="segmented">
          {papers.map(option => (
            <button key={option.name} type="button" aria-pressed={paper === option.name} onClick={() => {
              const size = pageFor(option.name, landscape)
              setPage(editor, { width: size.width, height: size.height })
            }}>{option.name}</button>
          ))}
        </span>
      </InspectorRow>
      {paper === undefined && <p className="inspector-note">This document uses its own paper size, {formatNumber(page.width / 1440)} × {formatNumber(page.height / 1440)} in.</p>}
      <InspectorRow label="Orientation">
        <span className="segmented">
          <button type="button" aria-pressed={!landscape} onClick={() => { if (landscape) setPage(editor, { width: page.height, height: page.width }) }}>Portrait</button>
          <button type="button" aria-pressed={landscape} onClick={() => { if (!landscape) setPage(editor, { width: page.height, height: page.width }) }}>Landscape</button>
        </span>
      </InspectorRow>
      {margin('marginTop', 'Top margin')}
      {margin('marginBottom', 'Bottom margin')}
      {margin('marginLeft', 'Left margin')}
      {margin('marginRight', 'Right margin')}
      <label className="sumi-check-row">
        <input type="checkbox" checked={page.pageNumbers} onChange={event => setPage(editor, { pageNumbers: event.currentTarget.checked })} />
        Page numbers in the footer
      </label>
      <p className="inspector-note">Used when printing and saved in Word files.</p>
    </InspectorSection>
  )
}

/** Detail that does not need to be one click away: paragraph and table settings, page setup, counts and the outline. */
export function DocInspector({ editor }: { readonly editor: Editor }) {
  const { counts, headings } = useEditorState({
    editor,
    selector: ({ editor: current }) => ({ counts: countText(current.state.doc), headings: outline(current.state.doc) }),
  })

  return (
    <Inspector label="Document inspector">
      <ParagraphSection editor={editor} />
      <TableSection editor={editor} />
      <PageSection editor={editor} />

      <InspectorSection title="Document">
        <InspectorRow label="Words"><span>{counts.words.toLocaleString()}</span></InspectorRow>
        <InspectorRow label="Characters"><span>{counts.characters.toLocaleString()}</span></InspectorRow>
      </InspectorSection>

      <InspectorSection title="Outline">
        {headings.length === 0
          ? <p className="inspector-note">Headings appear here. Choose Title or Heading 1 to 3 from the style menu.</p>
          : (
            <ul className="inspector-list sumi-outline">
              {headings.map(heading => (
                <li key={heading.position} data-level={heading.level}>
                  <button
                    type="button"
                    className="button is-quiet"
                    onClick={() => editor.chain().focus().setTextSelection(heading.position).scrollIntoView().run()}
                  >
                    {heading.text === '' ? 'Untitled heading' : heading.text}
                  </button>
                </li>
              ))}
            </ul>
          )}
      </InspectorSection>
    </Inspector>
  )
}
