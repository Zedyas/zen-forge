/**
 * The page Sumi shows, writes to .docx and will print: US Letter where it is the local standard,
 * A4 elsewhere, always with 1-inch margins. Sizes are in twips (1/1440 inch), Word's own unit.
 */
export interface PageFormat {
  readonly name: 'Letter' | 'A4'
  readonly width: number
  readonly height: number
  readonly margin: number
}

export const letterPage: PageFormat = { name: 'Letter', width: 12_240, height: 15_840, margin: 1_440 }
export const a4Page: PageFormat = { name: 'A4', width: 11_906, height: 16_838, margin: 1_440 }

/** Regions that use US Letter paper; everywhere else uses A4. */
const letterRegions = new Set(['US', 'CA', 'MX', 'PH', 'CL', 'CO', 'VE', 'PR', 'GT', 'CR', 'PA', 'DO', 'SV'])

export function localPageFormat(locale: string = navigator.language): PageFormat {
  try {
    const region = new Intl.Locale(locale).maximize().region
    return region !== undefined && letterRegions.has(region) ? letterPage : a4Page
  } catch {
    return a4Page
  }
}

export const twipsPerPixel = 15

/** The width text can use, in CSS pixels (96 per inch). */
export function contentWidthPixels(page: PageFormat): number {
  return (page.width - 2 * page.margin) / twipsPerPixel
}

export function pageSizeLabel(page: PageFormat): string {
  return page.name === 'Letter' ? 'Letter, 8.5 × 11 in' : 'A4, 210 × 297 mm'
}
