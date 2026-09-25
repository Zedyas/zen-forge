import { FolderSearch, Minus, Plus, Trash2 } from 'lucide-react'
import type { OpenDocument } from '../../app/documents-store'
import { platformClient } from '../../services/platform/client'
import { Inspector, InspectorRow, InspectorSection } from '../../ui/Inspector'
import { Tip } from '../../ui/Tip'
import { findMarkup, removeMarkup, updateMarkup } from './markup-actions'
import type { Markup } from './model'
import { commitPdf, pageSize, type ReadyPdf } from './pdf-store'
import { removeSignature, useSignatureStore } from './signatures'
import { setTool, useToolStore } from './tool-store'

const markupNames: Record<Markup['kind'], string> = {
  text: 'Text',
  image: 'Signature or image',
  rect: 'Shape',
  ink: 'Drawing',
  redact: 'Redaction',
}

function describeMarkup(markup: Markup): string {
  if (markup.kind !== 'rect') return markupNames[markup.kind]
  if (markup.fill === '#ffffff' && markup.opacity === 1) return 'White-out'
  return markup.stroke !== undefined ? 'Rectangle' : 'Highlight'
}

function points(value: number): string {
  return `${Math.round(value)} pt`
}

function SelectedMarkup({ documentId, document }: { readonly documentId: string; readonly document: ReadyPdf }) {
  const selection = useToolStore(state => state.selection)
  if (selection === undefined) return null
  const markup = findMarkup(document.present, selection.pageKey, selection.markupId)
  if (markup === undefined) return null
  const update = (next: Markup): void => updateMarkup(documentId, selection.pageKey, selection.markupId, next)

  return (
    <InspectorSection title={describeMarkup(markup)}>
      {markup.kind === 'text' && (
        <InspectorRow label="Size">
          <span className="pdf-stepper">
            <Tip label="Smaller text">
              <button type="button" className="tool" aria-label="Smaller text" onClick={() => update({ ...markup, size: Math.max(6, markup.size - 1) })}><Minus aria-hidden="true" size={13} /></button>
            </Tip>
            <span>{markup.size} pt</span>
            <Tip label="Larger text">
              <button type="button" className="tool" aria-label="Larger text" onClick={() => update({ ...markup, size: Math.min(96, markup.size + 1) })}><Plus aria-hidden="true" size={13} /></button>
            </Tip>
          </span>
        </InspectorRow>
      )}
      {(markup.kind === 'image' || markup.kind === 'rect' || markup.kind === 'redact') && (
        <InspectorRow label="Size"><span>{points(markup.width)} × {points(markup.height)}</span></InspectorRow>
      )}
      {markup.kind === 'redact' && (
        <p className="inspector-note">Removed permanently when you save. The page becomes an image.</p>
      )}
      <button type="button" className="button is-quiet pdf-inspector-action" onClick={() => removeMarkup(documentId, selection.pageKey, selection.markupId)}>
        <Trash2 aria-hidden="true" size={14} />Delete
      </button>
    </InspectorSection>
  )
}

function Signatures() {
  const signatures = useSignatureStore(state => state.signatures)
  const armed = useToolStore(state => state.armedSignatureId)
  return (
    <InspectorSection title="Signatures">
      {signatures.length === 0 && <p className="inspector-note">Draw your signature once and place it on any PDF.</p>}
      <ul className="pdf-signature-list">
        {signatures.map(signature => (
          <li key={signature.id}>
            <button
              type="button"
              className={`pdf-signature-choice${armed === signature.id ? ' is-armed' : ''}`}
              aria-label="Place this signature"
              onClick={() => {
                setTool('signature')
                useToolStore.setState({ armedSignatureId: signature.id })
              }}
            >
              <img src={signature.dataUrl} alt="" />
            </button>
            <Tip label="Delete signature">
              <button type="button" className="tool" aria-label="Delete signature" onClick={() => removeSignature(signature.id)}>
                <Trash2 aria-hidden="true" size={14} />
              </button>
            </Tip>
          </li>
        ))}
      </ul>
      <button type="button" className="button is-quiet pdf-inspector-action" onClick={() => useToolStore.setState({ signatureDialogOpen: true })}>
        <Plus aria-hidden="true" size={14} />Add signature…
      </button>
    </InspectorSection>
  )
}

function PageDetails({ document }: { readonly document: ReadyPdf }) {
  const item = document.present.pages[document.currentPage]
  if (item === undefined) return null
  const size = pageSize(document, item)
  return (
    <InspectorSection title={`Page ${document.currentPage + 1}`}>
      <InspectorRow label="Size"><span>{points(size.width)} × {points(size.height)}</span></InspectorRow>
      <InspectorRow label="Inches"><span>{(size.width / 72).toFixed(2)} × {(size.height / 72).toFixed(2)}</span></InspectorRow>
      <InspectorRow label="Rotation"><span>{item.rotation}°</span></InspectorRow>
    </InspectorSection>
  )
}

function FormDetails({ documentId, document }: { readonly documentId: string; readonly document: ReadyPdf }) {
  const fillable = document.fields.filter(field => field.type !== 'button' && field.type !== 'signature' && field.type !== 'unknown')
  if (fillable.length === 0) return null
  const filled = fillable.filter(field => {
    const value = document.present.formValues[field.name] ?? field.value
    return value !== '' && value !== false
  }).length
  return (
    <InspectorSection title="Form">
      <InspectorRow label="Filled"><span>{filled} of {fillable.length}</span></InspectorRow>
      <label className="pdf-check-row">
        <input
          type="checkbox"
          checked={document.present.flattenForm}
          onChange={event => {
            const flattenForm = event.currentTarget.checked
            commitPdf(documentId, snapshot => ({ ...snapshot, flattenForm }))
          }}
        />
        Flatten form on save
      </label>
      <p className="inspector-note">Flattening turns the fields into plain text so the values cannot be changed later.</p>
    </InspectorSection>
  )
}

export function PdfInspector({ documentId, document, openDocument }: {
  readonly documentId: string
  readonly document: ReadyPdf
  readonly openDocument: OpenDocument
}) {
  return (
    <Inspector label="PDF inspector">
      <SelectedMarkup documentId={documentId} document={document} />
      <Signatures />
      <PageDetails document={document} />
      <FormDetails documentId={documentId} document={document} />
      <InspectorSection title="Document">
        <InspectorRow label="Pages"><span>{document.present.pages.length}</span></InspectorRow>
        <InspectorRow label="File"><span className="pdf-file-name">{openDocument.name}.{openDocument.extension}</span></InspectorRow>
        {openDocument.path !== undefined && (
          <button type="button" className="button is-quiet pdf-inspector-action" onClick={() => {
            if (openDocument.path !== undefined) void platformClient.revealInFinder(openDocument.path)
          }}>
            <FolderSearch aria-hidden="true" size={14} />Show in Finder
          </button>
        )}
      </InspectorSection>
    </Inspector>
  )
}
