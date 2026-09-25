import { app } from 'electron'
import { paperForRegion, type Paper } from '../../src/shared/print'
import { handle } from './security'

function defaultPaper(): Paper {
  return paperForRegion(app.getLocaleCountryCode())
}

/**
 * Printing for the renderer's print service, which mounts a printout into its page and then asks
 * for that page to be printed. No channel takes arguments: the sender's own page is what prints,
 * and the page's CSS `@page` rules set the page sizes.
 */
export function registerPrintIpc(): void {
  handle('print:paper', () => defaultPaper())

  // The macOS print dialog; its PDF menu also offers Save as PDF. Resolves false when cancelled.
  handle('print:print', event => new Promise<boolean>((resolve, reject) => {
    event.sender.print({ printBackground: true, pageSize: defaultPaper().name }, (printed, failureReason) => {
      if (printed || failureReason === 'Print job canceled') resolve(printed)
      else reject(new Error(failureReason))
    })
  }))

  handle('print:to-pdf', event =>
    event.sender.printToPDF({ printBackground: true, preferCSSPageSize: true, pageSize: defaultPaper().name }),
  )
}
