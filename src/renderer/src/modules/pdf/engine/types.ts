import type { ImportFindingInput } from '@shared/fidelity'

/** `#rrggbb` or `#rgb`. */
export type PdfColor = `#${string}`

export interface PdfPoint {
  readonly x: number
  readonly y: number
}

/** Clockwise page rotation, in degrees. */
export type PageRotation = 0 | 90 | 180 | 270

/**
 * Every coordinate is in PDF points with the origin at the top-left of the page
 * *as displayed* — after the page's `/Rotate` and any user rotation have been
 * applied, with y growing downwards.
 */
export type PdfEdit =
  | {
      readonly kind: 'text'
      readonly x: number
      readonly y: number
      readonly text: string
      readonly size: number
      readonly color: PdfColor
    }
  | {
      readonly kind: 'image'
      readonly x: number
      readonly y: number
      readonly width: number
      readonly height: number
      readonly png: Uint8Array
    }
  | {
      readonly kind: 'rect'
      readonly x: number
      readonly y: number
      readonly width: number
      readonly height: number
      readonly fill?: PdfColor
      readonly stroke?: PdfColor
      readonly opacity: number
    }
  | {
      readonly kind: 'ink'
      readonly points: readonly PdfPoint[]
      readonly width: number
      readonly color: PdfColor
    }

export interface PageRef {
  /** Index into `SavePdfInput.sources`. */
  readonly source: number
  /** Page index within that source. */
  readonly index: number
  /** Additional clockwise rotation the user applied, on top of the page's own `/Rotate`. */
  readonly rotation: PageRotation
  readonly edits: readonly PdfEdit[]
}

export type FormFieldType =
  | 'text'
  | 'checkbox'
  | 'radio'
  | 'dropdown'
  | 'list'
  | 'button'
  | 'signature'
  | 'unknown'

/** Widget rectangle in displayed top-left coordinates. */
export interface FormFieldWidget {
  readonly page: number
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

export interface FormField {
  readonly name: string
  readonly type: FormFieldType
  /** Checkboxes report a boolean; every other type reports its first selected value. */
  readonly value: string | boolean
  readonly options?: readonly string[]
  readonly readOnly: boolean
  readonly multiline?: boolean
  readonly widgets: readonly FormFieldWidget[]
}

export interface SavePdfInput {
  readonly sources: readonly Uint8Array[]
  readonly pages: readonly PageRef[]
  readonly formValues?: Readonly<Record<string, string | boolean>>
  readonly flattenForm?: boolean
}

export interface PdfInspection {
  readonly pageCount: number
  readonly findings: readonly ImportFindingInput[]
  readonly hasForm: boolean
}

export interface ImageSource {
  readonly bytes: Uint8Array
  readonly type: 'png' | 'jpeg'
}

export interface RedactionRequest {
  readonly index: number
  /** Boxes in displayed page coordinates: points from the top-left of the page as shown. */
  readonly boxes: readonly { readonly x: number; readonly y: number; readonly width: number; readonly height: number }[]
}

export interface RedactOptions {
  /** Also remove metadata, bookmarks, attachments, comments, hidden layers and scripts, like Acrobat's Remove Hidden Information. */
  readonly removeHiddenInformation: boolean
}
