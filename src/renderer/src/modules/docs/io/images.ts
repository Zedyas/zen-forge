/** Image bytes as the editor holds them (data URLs) and the pixel size Word needs to place them. */

export interface PixelSize {
  readonly width: number
  readonly height: number
}

export function toDataUrl(bytes: Uint8Array, mimeType: string): string {
  let binary = ''
  // Chunked, because spreading a large array into one call overflows the argument limit.
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000))
  }
  return `data:${mimeType};base64,${btoa(binary)}`
}

export function fromDataUrl(src: string): { readonly mimeType: string; readonly bytes: Uint8Array } | undefined {
  const match = /^data:([^;,]+);base64,(.*)$/s.exec(src)
  if (match === null) return undefined
  const binary = atob(match[2] ?? '')
  return { mimeType: match[1] ?? '', bytes: Uint8Array.from(binary, character => character.charCodeAt(0)) }
}

export function mimeTypeForExtension(extension: string): string {
  return extension.toLowerCase() === 'png' ? 'image/png' : 'image/jpeg'
}

function bigEndian16(bytes: Uint8Array, at: number): number {
  return ((bytes[at] ?? 0) << 8) | (bytes[at + 1] ?? 0)
}

function littleEndian16(bytes: Uint8Array, at: number): number {
  return (bytes[at] ?? 0) | ((bytes[at + 1] ?? 0) << 8)
}

function jpegSize(bytes: Uint8Array): PixelSize | undefined {
  let at = 2
  while (at + 9 < bytes.length && bytes[at] === 0xff) {
    const marker = bytes[at + 1] ?? 0
    // Start-of-frame markers carry the size; C4, C8 and CC share the range but are not frames.
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return { height: bigEndian16(bytes, at + 5), width: bigEndian16(bytes, at + 7) }
    }
    at += 2 + bigEndian16(bytes, at + 2)
  }
  return undefined
}

/** Reads the pixel size from a PNG, JPEG, GIF or BMP header. */
export function imageSize(bytes: Uint8Array): PixelSize | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes.length >= 24) {
    return { width: view.getUint32(16), height: view.getUint32(20) }
  }
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return jpegSize(bytes)
  if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes.length >= 10) {
    return { width: littleEndian16(bytes, 6), height: littleEndian16(bytes, 8) }
  }
  if (bytes[0] === 0x42 && bytes[1] === 0x4d && bytes.length >= 26) {
    return { width: view.getInt32(18, true), height: Math.abs(view.getInt32(22, true)) }
  }
  return undefined
}
