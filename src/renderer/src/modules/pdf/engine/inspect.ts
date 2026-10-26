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
      suggestedAlternative: 'A file Zen Suzu saves has no password and no permission restrictions.',
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
    encrypted: doc.isEncrypted,
  }
}

/**
 * pdf-lib keeps an encrypted file's strings and streams encrypted while it writes new objects in
 * the clear, so a save would mix the two into a damaged file. A file with only permission
 * restrictions (an owner password, no password to open it) is decrypted by MuPDF into a plain PDF
 * first. A file that needs a password to open is refused.
 */
async function decrypt(bytes: Uint8Array): Promise<Uint8Array> {
  // Loaded on first use: the WebAssembly module is large and only restricted files and redaction need it.
  const mupdf = await import('mupdf')
  const document = new mupdf.PDFDocument(bytes)
  try {
    if (document.needsPassword()) {
      const error = new Error('This PDF is password-protected. Opening protected PDFs is not supported yet.')
      // pdf-actions explains this error by name, as it does pdf.js's own.
      error.name = 'PasswordException'
      throw error
    }
    return document.saveToBuffer('decrypt').asUint8Array().slice()
  } finally {
    document.destroy()
  }
}

/**
 * Inspects a PDF read from disk and gives the bytes to edit: the file itself, or for an encrypted
 * one its plain copy. The inspection is of the file as read, so it still lists the restrictions a
 * save removes.
 */
export async function preparePdf(bytes: Uint8Array): Promise<{ readonly bytes: Uint8Array; readonly inspection: PdfInspection }> {
  const inspection = await inspectPdf(bytes)
  return { bytes: inspection.encrypted ? await decrypt(bytes) : bytes, inspection }
}
