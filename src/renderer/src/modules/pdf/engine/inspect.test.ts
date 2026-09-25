import { PDFDocument } from '@cantoo/pdf-lib'
import { describe, expect, it } from 'vitest'
import { inspectPdf } from './inspect'

async function plain(): Promise<PDFDocument> {
  const doc = await PDFDocument.create()
  doc.addPage([400, 600])
  return doc
}

describe('inspectPdf', () => {
  it('reports nothing for a plain document', async () => {
    const report = await inspectPdf(await (await plain()).save())
    expect(report).toEqual({ pageCount: 1, findings: [], hasForm: false })
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
})
