import { toast } from 'sonner'
import { fileService } from '../../services/file/IpcFileService'
import { openFiles } from '../../app/document-actions'
import { imagesToPdf, inspectPdf, savePdf, type ImageSource, type PageRef } from './engine'

function report(error: unknown, title: string): void {
  toast.error(title, { description: error instanceof Error ? error.message : undefined })
}

async function writeAndOpen(bytes: Uint8Array, defaultName: string): Promise<void> {
  const path = await fileService.chooseSavePath({ defaultName, extensions: ['pdf'] })
  if (path === undefined) return
  await fileService.write(path, bytes)
  await openFiles([await fileService.describe(path)])
}

/** Chooses photos or scans and makes one PDF page per image, each page sized to its image. */
export async function createPdfFromImages(): Promise<void> {
  try {
    const files = await fileService.chooseFiles({ extensions: ['png', 'jpg', 'jpeg'], allowMultiple: true })
    if (files.length === 0) return
    const images: ImageSource[] = await Promise.all(files.map(async file => ({
      bytes: await fileService.read(file.path),
      type: file.extension === 'png' ? 'png' : 'jpeg',
    })))
    const bytes = await imagesToPdf(images)
    await writeAndOpen(bytes, files.length === 1 ? `${files[0]?.name.replace(/\.[^.]+$/, '') ?? 'Images'}.pdf` : 'Images.pdf')
  } catch (error) {
    report(error, 'Could not make a PDF from those images')
  }
}

/** Chooses two or more PDFs and joins all their pages, in the order chosen, into a new file. */
export async function combinePdfs(): Promise<void> {
  try {
    const files = await fileService.chooseFiles({ extensions: ['pdf'], allowMultiple: true })
    if (files.length === 0) return
    if (files.length === 1) {
      toast.info('Choose two or more PDFs to combine.', { description: 'Hold ⌘ to select several files.' })
      return
    }
    const sources = await Promise.all(files.map(file => fileService.read(file.path)))
    const counts = await Promise.all(sources.map(async source => (await inspectPdf(source)).pageCount))
    const pages: PageRef[] = counts.flatMap((count, source) =>
      Array.from({ length: count }, (_, index) => ({ source, index, rotation: 0, edits: [] })))
    await writeAndOpen(await savePdf({ sources, pages }), 'Combined.pdf')
  } catch (error) {
    report(error, 'Could not combine those PDFs')
  }
}
