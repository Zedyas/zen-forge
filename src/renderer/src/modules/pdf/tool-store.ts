import { create } from 'zustand'
import type { PdfColor, PdfPoint } from './engine'

export const toolIds = ['select', 'text', 'highlight', 'ink', 'rect', 'whiteout', 'signature', 'redact'] as const
export type Tool = (typeof toolIds)[number]

/** Single-key shortcuts: tooltips show them and the PDF tab listens for them. */
export const toolKeys: Readonly<Record<Tool, string>> = {
  select: 'V',
  text: 'T',
  highlight: 'H',
  ink: 'D',
  rect: 'R',
  whiteout: 'W',
  signature: 'S',
  redact: 'X',
}

export interface Selection {
  readonly pageKey: string
  readonly markupId: string
}

/** A text box being typed: either a new one at a point, or an existing markup. */
export interface TextDraft {
  readonly pageKey: string
  readonly markupId?: string
  readonly point: PdfPoint
}

interface ToolState {
  readonly tool: Tool
  readonly inkColor: PdfColor
  readonly highlightColor: PdfColor
  readonly textSize: number
  readonly selection?: Selection
  readonly textDraft?: TextDraft
  /** The saved signature the next click on a page places. */
  readonly armedSignatureId?: string
  readonly signatureDialogOpen: boolean
}

export const useToolStore = create<ToolState>(() => ({
  tool: 'select',
  inkColor: '#1b1b1b',
  highlightColor: '#ffd60a',
  textSize: 12,
  signatureDialogOpen: false,
}))

export function setTool(tool: Tool): void {
  useToolStore.setState(state => ({
    tool,
    selection: tool === 'select' ? state.selection : undefined,
    textDraft: undefined,
    armedSignatureId: tool === 'signature' ? state.armedSignatureId : undefined,
  }))
}

export function select(selection: Selection | undefined): void {
  useToolStore.setState({ selection })
}

export const inkColors: ReadonlyArray<{ readonly label: string; readonly value: PdfColor }> = [
  { label: 'Black', value: '#1b1b1b' },
  { label: 'Dark blue', value: '#1c3f94' },
  { label: 'Blue', value: '#2f6fdb' },
  { label: 'Red', value: '#c62828' },
  { label: 'Orange', value: '#d9730d' },
  { label: 'Green', value: '#1b7f3b' },
  { label: 'Purple', value: '#6b3fa0' },
  { label: 'Grey', value: '#6b7280' },
]

export const highlightColors: ReadonlyArray<{ readonly label: string; readonly value: PdfColor }> = [
  { label: 'Yellow', value: '#ffd60a' },
  { label: 'Green', value: '#7ee081' },
  { label: 'Blue', value: '#7cc4ff' },
  { label: 'Pink', value: '#ff8fc7' },
  { label: 'Orange', value: '#ffae42' },
  { label: 'Purple', value: '#c4a1ff' },
]
