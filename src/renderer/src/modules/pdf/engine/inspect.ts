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
      construct: 'Encryption and permission restrictions',
      severity: 'dropped',
      location: 'A saved copy is written unencrypted, without the original password or permissions.',
    })
  }

  if (form?.hasXFA() === true) {
    findings.push({
      construct: 'XFA form',
      severity: 'dropped',
      location: 'Dynamic XFA layouts cannot be rendered or filled here.',
      suggestedAlternative: 'Open in Adobe Acrobat',
    })
  }

  const signatures = form === undefined ? [] : form.getSignatureFields()
  if (signatures.length > 0) {
    findings.push({
      construct: 'Digital signatures',
      severity: 'dropped',
      location: `${signatures.length} signature field${signatures.length === 1 ? '' : 's'}; any save invalidates the signature.`,
    })
  }

  const attachments = doc.getAttachments()
  if (attachments.length > 0) {
    findings.push({
      construct: 'Embedded file attachments',
      severity: 'degraded',
      location: `${attachments.length} attached file${attachments.length === 1 ? '' : 's'} are kept in this document but lost if its pages are merged into another PDF.`,
    })
  }

  const scripts = doc.getDocumentJavaScripts()
  if (scripts.length > 0) {
    findings.push({
      construct: 'JavaScript actions',
      severity: 'degraded',
      location: 'Scripts are never run here, and are lost if these pages are merged into another PDF or the form is flattened.',
    })
  }

  if (doc.catalog.get(PDFName.of('StructTreeRoot')) !== undefined) {
    findings.push({
      construct: 'Tagged reading order',
      severity: 'degraded',
      location: 'The accessibility structure is not rebuilt when pages are reordered, removed, or edited.',
    })
  }

  return {
    pageCount: doc.getPageCount(),
    findings,
    hasForm: (form?.getFields().length ?? 0) > 0,
  }
}
