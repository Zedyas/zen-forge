/** Printer paper. Sizes are in millimetres, portrait. */
export interface Paper {
  readonly name: 'Letter' | 'A4'
  readonly width: number
  readonly height: number
}

export const letterPaper: Paper = { name: 'Letter', width: 215.9, height: 279.4 }
export const a4Paper: Paper = { name: 'A4', width: 210, height: 297 }

/** Regions that use US Letter, from the Unicode CLDR paper size data. Everywhere else uses A4. */
const letterRegions: ReadonlySet<string> = new Set(['BZ', 'CA', 'CL', 'CO', 'CR', 'GT', 'MX', 'NI', 'PA', 'PH', 'PR', 'SV', 'US', 'VE'])

/** The paper printouts are laid out for, from a region code such as `US` (`app.getLocaleCountryCode()`). */
export function paperForRegion(countryCode: string): Paper {
  return letterRegions.has(countryCode.toUpperCase()) ? letterPaper : a4Paper
}
