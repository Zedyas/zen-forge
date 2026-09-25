import { unzipSync } from 'fflate'

/**
 * Office files (.xlsx, .docx, .pptx, .ods, .numbers) are zip packages, and a package of a few
 * kilobytes can unpack to gigabytes. Its directory lists every entry's unpacked size, so the total
 * is checked before anything is inflated; a file past the limit is refused instead of exhausting
 * memory and taking the window, with its unsaved tabs, down with it.
 */
const maximumUnpackedBytes = 1024 ** 3
const maximumEntries = 20_000

/** Throws a plain message when `bytes` is a zip whose contents are too large to open safely. */
export function assertZipFitsInMemory(bytes: Uint8Array): void {
  // Not a zip (an .xls, say): nothing to check.
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) return
  let total = 0
  let entries = 0
  unzipSync(bytes, {
    filter: file => {
      total += file.originalSize
      entries += 1
      // Nothing is inflated: the filter only reads the directory.
      return false
    },
  })
  if (total > maximumUnpackedBytes || entries > maximumEntries) {
    throw new Error('This file unpacks to more than 1 GB, which is too large to open safely.')
  }
}
