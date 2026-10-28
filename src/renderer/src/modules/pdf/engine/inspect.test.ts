import { PDFDocument, StandardFonts } from '@cantoo/pdf-lib'
import { describe, expect, it } from 'vitest'
import { inspectPdf, preparePdf } from './inspect'
import { savePdf } from './save'
import { extractPageTexts } from './test-support'

async function plain(): Promise<PDFDocument> {
  const doc = await PDFDocument.create()
  doc.addPage([400, 600])
  return doc
}

describe('inspectPdf', () => {
  it('reports nothing for a plain document', async () => {
    const report = await inspectPdf(await (await plain()).save())
    expect(report).toEqual({ pageCount: 1, findings: [], hasForm: false, encrypted: false })
  })

  it('reports encryption as dropped, because a save writes an unencrypted copy', async () => {
    const doc = await plain()
    doc.encrypt({ ownerPassword: 'owner', permissions: { printing: false } })
    const report = await inspectPdf(await doc.save())

    expect(report.findings.map(finding => [finding.construct, finding.severity])).toEqual([
      ['Password and permission restrictions', 'dropped'],
    ])
  })

  it('reports attachments and JavaScript as degraded, since merging loses them', async () => {
    const doc = await plain()
    await doc.attach(new Uint8Array([1, 2, 3]), 'receipt.txt', { mimeType: 'text/plain' })
    doc.addJavaScript('greet', 'app.alert("hi")')
    const report = await inspectPdf(await doc.save())

    expect(report.findings.map(finding => [finding.construct, finding.severity])).toEqual([
      ['Embedded file attachments', 'degraded'],
      ['JavaScript actions', 'degraded'],
    ])
  })

  it('flags a document that has form fields', async () => {
    const doc = await plain()
    doc.getForm().createTextField('name').addToPage(doc.getPages()[0], { x: 10, y: 10 })
    const report = await inspectPdf(await doc.save())

    expect(report.hasForm).toBe(true)
    expect(report.findings).toEqual([])
  })

  it('edits an owner-restricted PDF as a plain copy and still reports the restrictions', async () => {
    const restricted = await encrypted('')
    const { bytes, inspection } = await preparePdf(restricted)

    expect(inspection.findings.map(finding => finding.construct)).toEqual(['Password and permission restrictions'])
    expect((await inspectPdf(bytes)).encrypted).toBe(false)
    // A page copied in from a second source kept its text only once the file is decrypted.
    const saved = await savePdf({ sources: [bytes, bytes], pages: [{ source: 0, index: 0, rotation: 0, edits: [] }, { source: 1, index: 0, rotation: 0, edits: [] }] })
    expect(await extractPageTexts(saved)).toEqual(['KEEP ME', 'KEEP ME'])
  })

  it('opens an owner-restricted PDF with an attachment and object streams', async () => {
    const { inspection } = await preparePdf(await encrypted('', { attachment: true }))

    expect(inspection.findings.map(finding => finding.construct)).toEqual(['Password and permission restrictions', 'Embedded file attachments'])
  })

  it('refuses a PDF that needs a password to open', async () => {
    await expect(preparePdf(await encrypted('user'))).rejects.toMatchObject({ name: 'PasswordException' })
  })
})

/** A one-page PDF encrypted by MuPDF with AES, its printing and copying restricted. */
async function encrypted(userPassword: string, { attachment = false } = {}): Promise<Uint8Array> {
  const doc = await plain()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  doc.getPage(0).drawText('KEEP ME', { x: 40, y: 500, size: 14, font })
  if (attachment) await doc.attach(new Uint8Array([1, 2, 3]), 'receipt.txt', { mimeType: 'text/plain' })
  const mupdf = await import('mupdf')
  const document = new mupdf.PDFDocument(await doc.save())
  try {
    const streams = attachment ? ',compress,objstms' : ''
    return document.saveToBuffer(`encrypt=aes-256,owner-password=owner,user-password=${userPassword},permissions=-3904${streams}`).asUint8Array().slice()
  } finally {
    document.destroy()
  }
}
