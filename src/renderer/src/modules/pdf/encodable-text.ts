import { toast } from 'sonner'

/** Characters outside WinAnsi, which Helvetica (used for added text and form appearances) cannot encode. */
const unencodable = /[^\n\u0020-\u007e\u00a0-\u00ff€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ]/g

/**
 * Removes what the saved PDF could not show, and says so. Left in, one such character makes
 * pdf-lib throw while writing the file.
 */
export function encodableText(text: string): string {
  const cleaned = text.replace(unencodable, '')
  if (cleaned !== text) {
    toast.info('Some characters were left out', {
      description: 'Text added to a PDF uses Helvetica, which covers Latin letters, digits and common symbols.',
    })
  }
  return cleaned
}
