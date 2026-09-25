import { create } from 'zustand'

export interface SavedSignature {
  readonly id: string
  /** PNG, cropped to the ink, transparent background. */
  readonly dataUrl: string
  /** Natural size in CSS pixels; only the ratio matters when placing. */
  readonly width: number
  readonly height: number
}

const storageKey = 'signatures'

function isSignature(value: unknown): value is SavedSignature {
  return typeof value === 'object' && value !== null
    && 'id' in value && typeof value.id === 'string'
    && 'dataUrl' in value && typeof value.dataUrl === 'string' && value.dataUrl.startsWith('data:image/png;base64,')
    && 'width' in value && typeof value.width === 'number'
    && 'height' in value && typeof value.height === 'number'
}

function readSignatures(): readonly SavedSignature[] {
  try {
    const parsed: unknown = JSON.parse(localStorage.getItem(storageKey) ?? '[]')
    return Array.isArray(parsed) ? parsed.filter(isSignature) : []
  } catch {
    return []
  }
}

function write(signatures: readonly SavedSignature[]): void {
  try {
    localStorage.setItem(storageKey, JSON.stringify(signatures))
  } catch {
    // Storage full or unavailable: the signature still works for this session.
  }
}

/** Signatures drawn once and reused, kept on this Mac only. */
export const useSignatureStore = create<{ readonly signatures: readonly SavedSignature[] }>(() => ({
  signatures: readSignatures(),
}))

export function addSignature(signature: SavedSignature): void {
  const signatures = [signature, ...useSignatureStore.getState().signatures]
  useSignatureStore.setState({ signatures })
  write(signatures)
}

export function removeSignature(id: string): void {
  const signatures = useSignatureStore.getState().signatures.filter(signature => signature.id !== id)
  useSignatureStore.setState({ signatures })
  write(signatures)
}

export function dataUrlBytes(dataUrl: string): Uint8Array {
  const base64 = dataUrl.slice(dataUrl.indexOf(',') + 1)
  return Uint8Array.from(atob(base64), character => character.charCodeAt(0))
}

const imageUrls = new WeakMap<Uint8Array, string>()

/**
 * A data URL for PNG bytes, encoded once per byte array. Unlike an object URL it needs no revoking:
 * it is collected together with the bytes when the markup is gone.
 */
export function pngUrl(png: Uint8Array): string {
  const cached = imageUrls.get(png)
  if (cached !== undefined) return cached
  let binary = ''
  // Chunked, because spreading a large array into one call exceeds the argument limit.
  for (let offset = 0; offset < png.length; offset += 0x8000) {
    binary += String.fromCharCode(...png.subarray(offset, offset + 0x8000))
  }
  const url = `data:image/png;base64,${btoa(binary)}`
  imageUrls.set(png, url)
  return url
}
