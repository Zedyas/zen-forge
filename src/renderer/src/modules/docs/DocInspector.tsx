import type { Editor } from '@tiptap/core'
import { useEditorState } from '@tiptap/react'
import { Inspector, InspectorRow, InspectorSection } from '../../ui/Inspector'
import { page } from './doc-documents'
import { pageSizeLabel } from './page'
import { countText, outline } from './text-stats'

/** Detail that does not need to be one click away: counts, page setup and the outline. */
export function DocInspector({ editor }: { readonly editor: Editor }) {
  const { counts, headings } = useEditorState({
    editor,
    selector: ({ editor: current }) => ({ counts: countText(current.state.doc), headings: outline(current.state.doc) }),
  })

  return (
    <Inspector label="Document inspector">
      <InspectorSection title="Document">
        <InspectorRow label="Words"><span>{counts.words.toLocaleString()}</span></InspectorRow>
        <InspectorRow label="Characters"><span>{counts.characters.toLocaleString()}</span></InspectorRow>
      </InspectorSection>

      <InspectorSection title="Page">
        <InspectorRow label="Size"><span>{pageSizeLabel(page)}</span></InspectorRow>
        <InspectorRow label="Margins"><span>1 in</span></InspectorRow>
        <InspectorRow label="Text"><span>Arial, 11 pt</span></InspectorRow>
        <p className="inspector-note">Pages follow your region. Word files are saved with this page size.</p>
      </InspectorSection>

      <InspectorSection title="Outline">
        {headings.length === 0
          ? <p className="inspector-note">Headings appear here. Choose Heading 1 to 3 from the style menu.</p>
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
