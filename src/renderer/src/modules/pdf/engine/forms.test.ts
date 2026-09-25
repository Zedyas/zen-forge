import { PDFDocument } from '@cantoo/pdf-lib'
import { describe, expect, it } from 'vitest'
import { listFormFields } from './forms'
import { savePdf } from './save'
import { extractPageTexts } from './test-support'

async function formDocument(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([400, 600])
  const form = doc.getForm()
  form.createTextField('applicantName').addToPage(page, { x: 50, y: 500, width: 200, height: 20, borderWidth: 0 })
  form.createCheckBox('agree').addToPage(page, { x: 50, y: 460, width: 16, height: 16, borderWidth: 0 })
  return doc.save()
}

const onePage = [{ source: 0, index: 0, rotation: 0, edits: [] }] as const

describe('AcroForm fields', () => {
  it('reports fields with values and displayed widget rectangles', async () => {
    const filled = await savePdf({
      sources: [await formDocument()],
      pages: onePage,
      formValues: { applicantName: 'Jane Quill', agree: true },
    })

    const fields = await listFormFields(filled)
    expect(fields.map(field => [field.name, field.type, field.value])).toEqual([
      ['applicantName', 'text', 'Jane Quill'],
      ['agree', 'checkbox', true],
    ])
    expect(fields[0].widgets).toEqual([{ page: 0, x: 50, y: 80, width: 200, height: 20 }])
    expect(fields[0].readOnly).toBe(false)
    expect(fields[0].multiline).toBe(false)
  })

  it('removes the fields and bakes in their text when flattening', async () => {
    const flattened = await savePdf({
      sources: [await formDocument()],
      pages: onePage,
      formValues: { applicantName: 'Jane Quill', agree: true },
      flattenForm: true,
    })

    expect(await listFormFields(flattened)).toEqual([])
    expect((await extractPageTexts(flattened))[0]).toContain('Jane Quill')
  })
})
