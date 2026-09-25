import { PDFDocument, PDFName } from '@cantoo/pdf-lib'
import type { ImportFindingInput } from '@shared/fidelity'
import type { PdfInspection } from './types'

export async function loadDocument(bytes: Uint8Array): Promise<PDFDocument> {
  return PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false })
}

/**
 * Reports only what a save would change. Attachments and document JavaScript
 * survive a save of this file on its own, because the editor rewrites the
 * primary document in place; they are listed as degraded because they are lost
 * when these pages are merged into a different PDF.
 */
export async function inspectPdf(bytes: Uint8Array): Promise<PdfInspection> {
  const doc = await loadDocument(bytes)
  const findings: ImportFindingInput[] = []
  const acroForm = doc.catalog.AcroForm()
  const form = acroForm === undefined ? undefined : doc.getForm()

  if (doc.isEncrypted) {
    findings.push({
      construct: 'Password and permission restrictions',
      severity: 'dropped',
      suggestedAlternative: 'A file Zendo saves has no password and no permission restrictions.',
    })
  }

  if (form?.hasXFA() === true) {
    findings.push({
      construct: 'XFA form',
      severity: 'dropped',
      suggestedAlternative: 'Its fields can’t be shown or filled here. Fill it in Adobe Acrobat instead.',
    })
  }

  const signatures = form === undefined ? [] : form.getSignatureFields()
  if (signatures.length > 0) {
    findings.push({
      construct: 'Digital signatures',
      severity: 'dropped',
      location: `${signatures.length} signature field${signatures.length === 1 ? '' : 's'}`,
      suggestedAlternative: 'Saving any change makes the signatures invalid. Save a copy to keep the signed original.',
    })
  }

  const attachments = doc.getAttachments()
  if (attachments.length > 0) {
    findings.push({
      construct: 'Embedded file attachments',
      severity: 'degraded',
      location: `${attachments.length} attached file${attachments.length === 1 ? '' : 's'}`,
      suggestedAlternative: 'Kept when you save this PDF; not carried over when its pages are combined into another PDF.',
    })
  }

  const scripts = doc.getDocumentJavaScripts()
  if (scripts.length > 0) {
    findings.push({
      construct: 'JavaScript actions',
      severity: 'degraded',
      suggestedAlternative: 'Never run here. Kept when you save; removed when the pages are combined into another PDF or the form is flattened.',
    })
  }

  if (doc.catalog.get(PDFName.of('StructTreeRoot')) !== undefined) {
    findings.push({
      construct: 'Tagged reading order',
      severity: 'degraded',
      suggestedAlternative: 'Kept, but not updated when pages are reordered, removed or edited, so screen readers may read them out of order.',
    })
  }

  return {
    pageCount: doc.getPageCount(),
    findings,
    hasForm: (form?.getFields().length ?? 0) > 0,
  }
}
