import { useEffect, useRef, useState } from 'react'
import { encodableText } from './encodable-text'
import type { FormField, FormFieldWidget, PageRotation } from './engine'
import { rotateBox, type Size } from './model'
import { commitPdf } from './pdf-store'

interface FormLayerProps {
  readonly documentId: string
  /** Fields with at least one widget on this page. */
  readonly fields: readonly FormField[]
  /** The source page index these widgets belong to. */
  readonly pageIndex: number
  /** The user's rotation of the page; widgets are turned with it. */
  readonly rotation: PageRotation
  /** The page's displayed size before the user's rotation, which widget rectangles are measured on. */
  readonly unrotatedSize: Size
  readonly values: Readonly<Record<string, string | boolean>>
  readonly zoom: number
}

function setValue(documentId: string, name: string, value: string | boolean): void {
  commitPdf(documentId, snapshot => snapshot.formValues[name] === value
    ? snapshot
    : { ...snapshot, formValues: { ...snapshot.formValues, [name]: value } })
}

function widgetStyle(widget: FormFieldWidget, zoom: number) {
  return {
    left: widget.x * zoom,
    top: widget.y * zoom,
    width: widget.width * zoom,
    height: widget.height * zoom,
    fontSize: Math.max(8, Math.min(widget.height * 0.62, 12)) * zoom,
  }
}

/**
 * A text field keeps typing local and records one undo step when it loses focus, or when it
 * unmounts mid-edit (switching tabs from the keyboard removes it without a blur).
 */
function TextField({ documentId, field, widget, value, zoom }: {
  readonly documentId: string
  readonly field: FormField
  readonly widget: FormFieldWidget
  readonly value: string
  readonly zoom: number
}) {
  const [draft, setDraft] = useState(value)
  const [editing, setEditing] = useState(false)
  /** The value when editing started and the text typed since; `undefined` while not editing. */
  const edit = useRef<{ readonly start: string; typed: string }>(undefined)
  const shown = editing ? draft : value

  const commit = (): void => {
    const current = edit.current
    edit.current = undefined
    if (current !== undefined && current.typed !== current.start) setValue(documentId, field.name, encodableText(current.typed))
  }

  useEffect(() => () => commit(), [])

  const props = {
    className: 'pdf-field',
    style: widgetStyle(widget, zoom),
    value: shown,
    disabled: field.readOnly,
    'aria-label': field.name,
    onFocus: () => {
      edit.current = { start: value, typed: value }
      setDraft(value)
      setEditing(true)
    },
    onChange: (event: { currentTarget: { value: string } }) => {
      if (edit.current !== undefined) edit.current.typed = event.currentTarget.value
      setDraft(event.currentTarget.value)
    },
    onBlur: () => {
      setEditing(false)
      commit()
    },
  }
  return field.multiline === true
    ? <textarea {...props} />
    : <input {...props} onKeyDown={event => { if (event.key === 'Enter') event.currentTarget.blur() }} />
}

/** Fillable form fields drawn as real inputs over their widgets. */
export function FormLayer({ documentId, fields, pageIndex, rotation, unrotatedSize, values, zoom }: FormLayerProps) {
  return (
    <div className="pdf-form-layer">
      {fields.flatMap(field => field.widgets.flatMap((placed, widgetIndex) => {
        if (placed.page !== pageIndex) return []
        const widget = rotateBox(placed, unrotatedSize, rotation)
        const value = values[field.name] ?? field.value
        const key = `${field.name}:${widgetIndex}`
        switch (field.type) {
          case 'text':
            return [<TextField key={key} documentId={documentId} field={field} widget={widget} value={String(value)} zoom={zoom} />]
          case 'checkbox':
            return [(
              <input
                key={key}
                type="checkbox"
                className="pdf-field pdf-check"
                style={widgetStyle(widget, zoom)}
                checked={value === true}
                disabled={field.readOnly}
                aria-label={field.name}
                onChange={event => setValue(documentId, field.name, event.currentTarget.checked)}
              />
            )]
          case 'radio': {
            const option = field.options?.[widgetIndex]
            if (option === undefined) return []
            return [(
              <input
                key={key}
                type="radio"
                name={`${documentId}:${field.name}`}
                className="pdf-field pdf-check"
                style={widgetStyle(widget, zoom)}
                checked={value === option}
                disabled={field.readOnly}
                aria-label={`${field.name}: ${option}`}
                onChange={() => setValue(documentId, field.name, option)}
              />
            )]
          }
          case 'dropdown':
          case 'list':
            return [(
              <select
                key={key}
                className="pdf-field"
                style={widgetStyle(widget, zoom)}
                value={String(value)}
                disabled={field.readOnly}
                aria-label={field.name}
                onChange={event => setValue(documentId, field.name, event.currentTarget.value)}
              >
                {value === '' && <option value="" />}
                {field.options?.map(option => <option key={option} value={option}>{option}</option>)}
              </select>
            )]
          default:
            return []
        }
      }))}
    </div>
  )
}
