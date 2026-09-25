import {
  PDFButton,
  PDFCheckBox,
  PDFDocument,
  PDFDropdown,
  PDFOptionList,
  PDFRadioGroup,
  PDFRef,
  PDFSignature,
  PDFTextField,
  type PDFField,
  type PDFPage,
  type PDFWidgetAnnotation,
} from '@cantoo/pdf-lib'
import { pageGeometry, rectToDisplayed } from './geometry'
import type { FormField, FormFieldType, FormFieldWidget } from './types'

function typeOf(field: PDFField): FormFieldType {
  if (field instanceof PDFTextField) return 'text'
  if (field instanceof PDFCheckBox) return 'checkbox'
  if (field instanceof PDFRadioGroup) return 'radio'
  if (field instanceof PDFDropdown) return 'dropdown'
  if (field instanceof PDFOptionList) return 'list'
  if (field instanceof PDFSignature) return 'signature'
  if (field instanceof PDFButton) return 'button'
  return 'unknown'
}

function valueOf(field: PDFField): string | boolean {
  if (field instanceof PDFTextField) return field.getText() ?? ''
  if (field instanceof PDFCheckBox) return field.isChecked()
  if (field instanceof PDFRadioGroup) return field.getSelected() ?? ''
  if (field instanceof PDFDropdown) return field.getSelected()[0] ?? ''
  if (field instanceof PDFOptionList) return field.getSelected()[0] ?? ''
  return ''
}

function optionsOf(field: PDFField): readonly string[] | undefined {
  if (field instanceof PDFRadioGroup) return field.getOptions()
  if (field instanceof PDFDropdown) return field.getOptions()
  if (field instanceof PDFOptionList) return field.getOptions()
  return undefined
}

/** Resolves a widget to a page index, by `/P` when present and by annotation scan otherwise. */
function buildPageLookup(pages: readonly PDFPage[]): (ref: PDFRef | undefined, widgetRef: PDFRef | undefined) => number | undefined {
  const byPageRef = new Map<string, number>()
  const byAnnotRef = new Map<string, number>()

  pages.forEach((page, index) => {
    byPageRef.set(page.ref.tag, index)
    const annots = page.node.Annots()
    if (annots === undefined) return
    for (let i = 0; i < annots.size(); i += 1) {
      const entry = annots.get(i)
      if (entry instanceof PDFRef) byAnnotRef.set(entry.tag, index)
    }
  })

  return (pageRef, widgetRef) => {
    const viaPage = pageRef === undefined ? undefined : byPageRef.get(pageRef.tag)
    if (viaPage !== undefined) return viaPage
    return widgetRef === undefined ? undefined : byAnnotRef.get(widgetRef.tag)
  }
}

export function readFormFields(doc: PDFDocument): FormField[] {
  if (doc.catalog.AcroForm() === undefined) return []
  const pages = doc.getPages()
  const geometries = pages.map(page => pageGeometry(page))
  const lookupPage = buildPageLookup(pages)

  return doc.getForm().getFields().map(field => {
    const widgets: FormFieldWidget[] = []
    for (const widget of field.acroField.getWidgets()) {
      const widgetRef = doc.context.getObjectRef(widget.dict)
      const pageIndex = lookupPage(widget.P(), widgetRef)
      if (pageIndex === undefined) continue
      const rect = rectToDisplayed(geometries[pageIndex], widget.getRectangle())
      widgets.push({ page: pageIndex, ...rect })
    }

    const type = typeOf(field)
    return {
      name: field.getName(),
      type,
      value: valueOf(field),
      options: optionsOf(field),
      readOnly: field.isReadOnly(),
      multiline: field instanceof PDFTextField ? field.isMultiline() : undefined,
      widgets,
    }
  })
}

/** Writes values onto the matching AcroForm fields; unknown names are ignored. */
export function applyFormValues(
  doc: PDFDocument,
  values: Readonly<Record<string, string | boolean>>,
): void {
  const form = doc.getForm()
  for (const [name, value] of Object.entries(values)) {
    const field = form.getFieldMaybe(name)
    if (field === undefined) continue

    if (field instanceof PDFCheckBox) {
      if (value === true || value === 'true') field.check()
      else field.uncheck()
    } else if (field instanceof PDFTextField) {
      field.setText(typeof value === 'boolean' ? String(value) : value)
    } else if (field instanceof PDFRadioGroup) {
      field.select(String(value))
    } else if (field instanceof PDFDropdown) {
      field.select(String(value))
    } else if (field instanceof PDFOptionList) {
      field.select(String(value))
    }
  }
}

/**
 * Removes the form widgets `doomed` picks, from their pages and their fields. A field left with
 * no widget goes whole, value included: nobody could see it, but its value would stay in the file.
 */
export function removeWidgets(doc: PDFDocument, doomed: (widget: PDFWidgetAnnotation, page: PDFPage | undefined) => boolean): void {
  if (doc.catalog.AcroForm() === undefined) return
  const form = doc.getForm()
  const pageOfWidget = new Map<string, PDFPage>()
  for (const page of doc.getPages()) {
    for (const ref of page.node.Annots()?.asArray() ?? []) if (ref instanceof PDFRef) pageOfWidget.set(ref.tag, page)
  }

  for (const field of form.getFields()) {
    const widgets = field.acroField.getWidgets()
    const refs = widgets.map(widget => doc.context.getObjectRef(widget.dict))
    const marked = widgets.map((widget, index) => {
      const ref = refs[index]
      return doomed(widget, ref === undefined ? undefined : pageOfWidget.get(ref.tag))
    })
    if (!marked.includes(true)) continue
    if (marked.every(Boolean)) {
      form.removeField(field)
      continue
    }
    for (let index = widgets.length - 1; index >= 0; index -= 1) {
      const ref = refs[index]
      if (!marked[index] || ref === undefined) continue
      pageOfWidget.get(ref.tag)?.node.removeAnnot(ref)
      field.acroField.removeWidget(index)
    }
  }
}

export async function listFormFields(bytes: Uint8Array): Promise<FormField[]> {
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false })
  return readFormFields(doc)
}
