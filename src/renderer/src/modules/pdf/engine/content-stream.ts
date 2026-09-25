/*
 * A small content-stream scanner. It finds where each operator and its operands sit, so a caller
 * can cut byte ranges out of a stream and leave every other byte as it was. It is not a parser:
 * operands are kept as the text they were written as, and nothing is interpreted except what is
 * needed to step over each token correctly:
 *   - literal strings `( … )`, which nest parentheses and escape them with a backslash,
 *   - hex strings `< … >`, dictionaries `<< … >>` and arrays `[ … ]`, which can hold any of these,
 *   - comments, from `%` to the end of the line,
 *   - inline images `BI … ID <binary data> EI`, whose data can hold any byte, including ones that
 *     look like the tokens above. See `inlineImageEnd` for how the end of the data is found.
 */

export interface Instruction {
  readonly operator: string
  /** Operands as written, except that names have their `#xx` escapes decoded. */
  readonly operands: readonly string[]
  /** Offset of the first operand, or of the operator when it has none. */
  readonly start: number
  /** Offset just past the operator; for an inline image (`BI`), just past its `EI`. */
  readonly end: number
}

const LF = 0x0a
const CR = 0x0d
const PAREN_OPEN = 0x28
const PAREN_CLOSE = 0x29
const SLASH = 0x2f
const LESS = 0x3c
const GREATER = 0x3e
const BRACKET_OPEN = 0x5b
const BRACKET_CLOSE = 0x5d
const BACKSLASH = 0x5c
const PERCENT = 0x25
const TILDE = 0x7e
const LETTER_E = 0x45
const LETTER_I = 0x49

const whitespace = new Set([0x00, 0x09, LF, 0x0c, CR, 0x20])
const delimiters = new Set(Array.from('()<>[]{}/%', character => character.charCodeAt(0)))

function isWhitespace(byte: number | undefined): boolean {
  return byte !== undefined && whitespace.has(byte)
}

/** Whether a token can end before this byte: the end of the data, whitespace or a delimiter. */
function isBoundary(byte: number | undefined): boolean {
  return byte === undefined || whitespace.has(byte) || delimiters.has(byte)
}

function text(bytes: Uint8Array, from: number, to: number): string {
  let result = ''
  for (let index = from; index < to; index += 1) result += String.fromCharCode(bytes[index] ?? 0)
  return result
}

/** Skips whitespace and comments. */
function skipSpace(bytes: Uint8Array, from: number): number {
  let pos = from
  while (pos < bytes.length) {
    if (isWhitespace(bytes[pos])) pos += 1
    else if (bytes[pos] === PERCENT) {
      while (pos < bytes.length && bytes[pos] !== LF && bytes[pos] !== CR) pos += 1
    } else break
  }
  return pos
}

/** The end of a run of regular characters: a number, a keyword, an operator or a name's body. */
function regularEnd(bytes: Uint8Array, from: number): number {
  let pos = from
  while (!isBoundary(bytes[pos])) pos += 1
  return pos
}

/** From an opening `(` to just past its matching `)`; a backslash escapes the byte after it. */
function stringEnd(bytes: Uint8Array, from: number): number {
  let depth = 0
  let pos = from
  while (pos < bytes.length) {
    const byte = bytes[pos]
    pos += 1
    if (byte === BACKSLASH) pos += 1
    else if (byte === PAREN_OPEN) depth += 1
    else if (byte === PAREN_CLOSE) {
      depth -= 1
      if (depth === 0) return pos
    }
  }
  return bytes.length
}

/** From an array's `[` or a dictionary's `<<` to just past its `]` or `>>`, stepping over every value inside. */
function containerEnd(bytes: Uint8Array, from: number, dictionary: boolean): number {
  let pos = from + (dictionary ? 2 : 1)
  while (true) {
    pos = skipSpace(bytes, pos)
    if (pos >= bytes.length) return bytes.length
    if (dictionary ? bytes[pos] === GREATER && bytes[pos + 1] === GREATER : bytes[pos] === BRACKET_CLOSE) {
      return pos + (dictionary ? 2 : 1)
    }
    pos = valueEnd(bytes, pos)
  }
}

/** Just past the value or word that starts at `from`. Always moves forward, even over a stray delimiter. */
function valueEnd(bytes: Uint8Array, from: number): number {
  const byte = bytes[from]
  if (byte === PAREN_OPEN) return stringEnd(bytes, from)
  if (byte === BRACKET_OPEN) return containerEnd(bytes, from, false)
  if (byte === LESS) {
    if (bytes[from + 1] === LESS) return containerEnd(bytes, from, true)
    const close = bytes.indexOf(GREATER, from)
    return close < 0 ? bytes.length : close + 1
  }
  if (byte === SLASH) return regularEnd(bytes, from + 1)
  if (isBoundary(byte)) return from + 1
  return regularEnd(bytes, from)
}

function decodeName(name: string): string {
  return name.replace(/#([0-9a-fA-F]{2})/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)))
}

/** Numbers, booleans and null are operands; every other bare word is an operator. */
function isOperandWord(word: string): boolean {
  return /^[+\-.\d]/.test(word) || word === 'true' || word === 'false' || word === 'null'
}

/** An `EI` token at `from`, after optional whitespace: returns the offset just past it. */
function endMarkerAt(bytes: Uint8Array, from: number): number | undefined {
  let pos = from
  while (isWhitespace(bytes[pos])) pos += 1
  return bytes[pos] === LETTER_E && bytes[pos + 1] === LETTER_I && isBoundary(bytes[pos + 2]) ? pos + 2 : undefined
}

/** How many bytes after a candidate `EI` must read as content-stream text for it to count as the end. */
const followingText = 16

/** The first `EI` token followed by text, as pdf.js finds it: compressed data rarely holds that by chance. */
function searchEndMarker(bytes: Uint8Array, from: number): number | undefined {
  for (let pos = bytes.indexOf(LETTER_E, from); pos >= 0; pos = bytes.indexOf(LETTER_E, pos + 1)) {
    if (bytes[pos + 1] !== LETTER_I || !isBoundary(bytes[pos + 2])) continue
    const following = bytes.subarray(pos + 2, pos + 2 + followingText)
    if (following.every(byte => isWhitespace(byte) || (byte >= 0x20 && byte < 0x7f))) return pos + 2
  }
  return undefined
}

/** The first name in a value such as `/Fl` or `[/AHx /Fl]`, without its slash. */
function firstName(value: string | undefined): string | undefined {
  return value === undefined ? undefined : /\/([^\s/[\]()<>{}%]+)/.exec(value)?.[1]
}

/** Colour components per pixel for the colour spaces an inline image can name without resources. */
const components: Readonly<Record<string, number>> = {
  G: 1, DeviceGray: 1, CalGray: 1, I: 1, Indexed: 1, RGB: 3, DeviceRGB: 3, CalRGB: 3, CMYK: 4, DeviceCMYK: 4,
}

/** The byte length of unfiltered image data, when the image's own entries determine it. */
function unfilteredLength(entries: ReadonlyMap<string, string>): number | undefined {
  if (firstName(entries.get('/F') ?? entries.get('/Filter')) !== undefined) return undefined
  const width = Number(entries.get('/W') ?? entries.get('/Width'))
  const height = Number(entries.get('/H') ?? entries.get('/Height'))
  const mask = (entries.get('/IM') ?? entries.get('/ImageMask')) === 'true'
  const bits = mask ? 1 : Number(entries.get('/BPC') ?? entries.get('/BitsPerComponent'))
  const space = firstName(entries.get('/CS') ?? entries.get('/ColorSpace'))
  const perPixel = mask ? 1 : space === undefined ? undefined : components[space]
  if (perPixel === undefined || ![width, height, bits].every(value => Number.isInteger(value) && value > 0)) return undefined
  return Math.ceil((width * perPixel * bits) / 8) * height
}

/**
 * Just past the `EI` of an inline image whose data starts at `dataStart`. In order of certainty:
 *   1. a declared data length (/L, from PDF 2.0),
 *   2. the length unfiltered data must have, from its width, height, bits and colour space,
 *   3. for ASCII filters, the first `EI` after the filter's end-of-data marker (`>` or `~>`),
 *      which cannot occur earlier in their data,
 *   4. otherwise the first `EI` followed by text (see `searchEndMarker`).
 * An image without an end cannot be stepped over, so scanning stops with an error rather than
 * reading its data as operators.
 */
function inlineImageEnd(bytes: Uint8Array, dataStart: number, entries: ReadonlyMap<string, string>): number {
  const declared = Number(entries.get('/L') ?? entries.get('/Length'))
  const length = Number.isInteger(declared) && declared >= 0 ? declared : unfilteredLength(entries)
  const exact = length === undefined ? undefined : endMarkerAt(bytes, dataStart + length)
  if (exact !== undefined) return exact

  const filter = firstName(entries.get('/F') ?? entries.get('/Filter'))
  const dataEnd = filter === 'AHx' || filter === 'ASCIIHexDecode'
    ? bytes.indexOf(GREATER, dataStart)
    : filter === 'A85' || filter === 'ASCII85Decode' ? bytes.indexOf(TILDE, dataStart) : dataStart
  const end = dataEnd < 0 ? undefined : searchEndMarker(bytes, dataEnd)
  if (end === undefined) throw new Error('A page contains an inline image whose end could not be found.')
  return end
}

/** From just after `BI`: reads the image's entries up to `ID`, then steps over its data and `EI`. */
function stepOverInlineImage(bytes: Uint8Array, from: number): number {
  const entries = new Map<string, string>()
  let key: string | undefined
  let pos = skipSpace(bytes, from)
  while (pos < bytes.length) {
    const tokenStart = pos
    pos = valueEnd(bytes, pos)
    const token = text(bytes, tokenStart, pos)
    // A single whitespace byte separates ID from the data.
    if (token === 'ID') return inlineImageEnd(bytes, isWhitespace(bytes[pos]) ? pos + 1 : pos, entries)
    const value = bytes[tokenStart] === SLASH ? decodeName(token) : token
    if (key === undefined) key = value
    else {
      entries.set(key, value)
      key = undefined
    }
    pos = skipSpace(bytes, pos)
  }
  throw new Error('A page contains an inline image whose end could not be found.')
}

/** Every operator in a content stream, in order, with the byte range it and its operands cover. */
export function scanContent(bytes: Uint8Array): Instruction[] {
  const instructions: Instruction[] = []
  let operands: string[] = []
  let start: number | undefined
  let pos = skipSpace(bytes, 0)
  while (pos < bytes.length) {
    const tokenStart = pos
    const first = bytes[pos]
    pos = valueEnd(bytes, pos)
    const token = text(bytes, tokenStart, pos)
    if (isBoundary(first) || isOperandWord(token)) {
      operands.push(first === SLASH ? decodeName(token) : token)
      start ??= tokenStart
    } else {
      if (token === 'BI') pos = stepOverInlineImage(bytes, pos)
      instructions.push({ operator: token, operands, start: start ?? tokenStart, end: pos })
      operands = []
      start = undefined
    }
    pos = skipSpace(bytes, pos)
  }
  return instructions
}
