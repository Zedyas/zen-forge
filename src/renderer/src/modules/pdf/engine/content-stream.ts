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
const SPACE = 0x20
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

/*
 * Where an inline image ends. Its data can hold any byte, `EI` included, so readers guess, and a
 * guess that differs from the viewer's puts different bytes in the image: hidden content could
 * then look like image data here and like operators there. So the end is found the way pdf.js
 * (the editor's viewer) finds it, step for step, and where the image declares or implies its
 * length and that points elsewhere, other readers would disagree, and scanning stops.
 */

/**
 * The operators pdf.js knows, with how many operands each takes; for colours, up to that many.
 * `null` marks the start of a longer word, which pdf.js reads on through (`BD` to `BDC`).
 */
const operandCounts: Readonly<Record<string, number | null>> = {
  w: 1, J: 1, j: 1, M: 1, d: 2, ri: 1, i: 1, gs: 1, q: 0, Q: 0, cm: 6, m: 2, l: 2, c: 6, v: 4, y: 4, h: 0, re: 4,
  S: 0, s: 0, f: 0, F: 0, 'f*': 0, B: 0, 'B*': 0, b: 0, 'b*': 0, n: 0, W: 0, 'W*': 0, BT: 0, ET: 0,
  Tc: 1, Tw: 1, Tz: 1, TL: 1, Tf: 2, Tr: 1, Ts: 1, Td: 2, TD: 2, Tm: 6, 'T*': 0, Tj: 1, TJ: 1, "'": 1, '"': 3,
  d0: 2, d1: 6, CS: 1, cs: 1, SC: 4, SCN: 33, sc: 4, scn: 33, G: 1, g: 1, RG: 3, rg: 3, K: 4, k: 4, sh: 1,
  BI: 0, ID: 0, EI: 1, Do: 1, MP: 1, DP: 2, BMC: 1, BDC: 2, EMC: 0, BX: 0, EX: 0,
  BM: null, BD: null, true: null, fa: null, fal: null, fals: null, false: null, nu: null, nul: null, null: null,
}
const variableOperands = new Set(['SC', 'SCN', 'sc', 'scn'])

const DIGIT_0 = 0x30
const DIGIT_9 = 0x39
const PLUS = 0x2b
const MINUS = 0x2d
const DOT = 0x2e

function isDigit(byte: number | undefined): boolean {
  return byte !== undefined && byte >= DIGIT_0 && byte <= DIGIT_9
}

/** Just past the number pdf.js reads at `from`, or undefined where it would fail to read one. */
function numberEnd(bytes: Uint8Array, from: number): number | undefined {
  let pos = from
  if (bytes[pos] === MINUS) pos += bytes[pos + 1] === MINUS ? 2 : 1
  else if (bytes[pos] === PLUS) pos += 1
  while (bytes[pos] === LF || bytes[pos] === CR) pos += 1
  let dot = bytes[pos] === DOT
  if (dot) pos += 1
  // pdf.js reads a sign with no digits as 0 when a space, string or the end follows, and fails otherwise.
  if (!isDigit(bytes[pos])) return isWhitespace(bytes[pos]) || bytes[pos] === PAREN_OPEN || bytes[pos] === LESS || pos >= bytes.length ? pos : undefined
  for (pos += 1; pos < bytes.length; pos += 1) {
    const byte = bytes[pos]
    if (byte === DOT && dot) break
    if (byte === DOT) dot = true
    else if (!isDigit(byte) && byte !== MINUS) break
  }
  return pos
}

/** Just past the word pdf.js reads at `from`: it stops where a word it knows would stop being one. */
function wordEnd(bytes: Uint8Array, from: number): number {
  let end = from + 1
  let known = text(bytes, from, end) in operandCounts
  while (!isBoundary(bytes[end])) {
    const longer = text(bytes, from, end + 1) in operandCounts
    if (known && !longer) break
    end += 1
    known = longer
  }
  return end
}

/**
 * Whether what follows an `EI` reads as content by pdf.js's test: operators it knows, until one
 * comes with as many operands as it takes. An unknown word, an array or dictionary (which pdf.js
 * does not expect here), or the end of the bytes given means no.
 */
function readsAsContent(bytes: Uint8Array): boolean {
  let operands = 0
  for (let pos = skipSpace(bytes, 0); pos < bytes.length; pos = skipSpace(bytes, pos)) {
    const byte = bytes[pos]
    if (byte === PAREN_OPEN || byte === SLASH || (byte === LESS && bytes[pos + 1] !== LESS)) {
      pos = valueEnd(bytes, pos)
      operands += 1
      continue
    }
    if (isBoundary(byte)) return false
    if (isDigit(byte) || byte === PLUS || byte === MINUS || byte === DOT) {
      const end = numberEnd(bytes, pos)
      if (end === undefined) return false
      pos = end
      operands += 1
      continue
    }
    const end = wordEnd(bytes, pos)
    const word = text(bytes, pos, end)
    pos = end
    if (word === 'true' || word === 'false' || word === 'null') {
      operands += 1
      continue
    }
    const takes = operandCounts[word]
    if (takes === undefined || takes === null) return false
    if (variableOperands.has(word) ? operands <= takes : operands === takes) return true
    operands = 0
  }
  return false
}

/** How many bytes after the space that follows an `EI` pdf.js checks are text, and how far it reads operators. */
const checkedText = 15
const lookAhead = 75

/**
 * Whether an `EI` at `pos` ends an image whose data has no end marker, by pdf.js's tests: a space
 * or line break after it, then text (a lone NUL allowed), then operators (see `readsAsContent`).
 */
function endsImage(bytes: Uint8Array, pos: number): boolean {
  if (bytes[pos] !== LETTER_E || bytes[pos + 1] !== LETTER_I) return false
  const space = bytes[pos + 2]
  if (space === undefined) return true
  if (space !== SPACE && space !== LF && space !== CR) return false
  const next = pos + 3
  const following = bytes.subarray(next, next + checkedText)
  for (let index = 0; index < following.length; index += 1) {
    const byte = following[index] ?? 0
    if (byte === 0 && following[index + 1] !== 0) continue
    if (byte !== LF && byte !== CR && (byte < SPACE || byte > 0x7f)) return false
  }
  return following.length === 0 || readsAsContent(bytes.subarray(next, next + lookAhead))
}

/** The first `EI` from `from` that ends the image, as pdf.js searches for it. */
function searchEndMarker(bytes: Uint8Array, from: number): number | undefined {
  for (let pos = bytes.indexOf(LETTER_E, from); pos >= 0; pos = bytes.indexOf(LETTER_E, pos + 1)) {
    if (endsImage(bytes, pos)) return pos
  }
  return undefined
}

/** JPEG markers followed by a segment with its length, which pdf.js steps over whole. */
const jpegSegments = new Set([0xc0, 0xc1, 0xc2, 0xc3, 0xc4, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcc, 0xcd, 0xce, 0xcf, 0xda, 0xdb, 0xdc, 0xdd, 0xde, 0xdf, 0xfe])
for (let marker = 0xe0; marker <= 0xef; marker += 1) jpegSegments.add(marker)

/** Just past a JPEG's end-of-image marker, walking its segments as pdf.js does. */
function jpegEnd(bytes: Uint8Array, from: number): number | undefined {
  let pos = from
  while (pos < bytes.length) {
    if (bytes[pos++] !== 0xff) continue
    const marker = bytes[pos++]
    if (marker === 0xd9) return pos
    if (marker === 0xff) pos -= 1
    else if (marker !== undefined && jpegSegments.has(marker)) {
      const length = ((bytes[pos] ?? 0) << 8) | (bytes[pos + 1] ?? 0)
      if (length > 2) pos += length
    }
  }
  return undefined
}

/** Just past ASCII85 data's end marker, `~>`, or a `~` that whitespace and `EI` follow, as pdf.js finds it. */
function ascii85End(bytes: Uint8Array, from: number): number | undefined {
  for (let pos = bytes.indexOf(TILDE, from); pos >= 0; pos = bytes.indexOf(TILDE, pos + 1)) {
    let next = pos + 1
    while (isWhitespace(bytes[next])) next += 1
    if (bytes[next] === GREATER) return next + 1
    if (next > pos + 1 && bytes[next] === LETTER_E && bytes[next + 1] === LETTER_I) return next
  }
  return undefined
}

/** Just past the end marker of JPEG, ASCII85 or ASCIIHex data; undefined for other data or without one. */
function endOfData(bytes: Uint8Array, dataStart: number, filter: string | undefined): number | undefined {
  if (filter === 'DCT' || filter === 'DCTDecode') return jpegEnd(bytes, dataStart)
  if (filter === 'A85' || filter === 'ASCII85Decode') return ascii85End(bytes, dataStart)
  if (filter !== 'AHx' && filter !== 'ASCIIHexDecode') return undefined
  const pos = bytes.indexOf(GREATER, dataStart)
  return pos < 0 ? undefined : pos + 1
}

/** The first `EI` from `from`, whatever follows it. */
function firstMarker(bytes: Uint8Array, from: number): number | undefined {
  for (let pos = bytes.indexOf(LETTER_E, from); pos >= 0; pos = bytes.indexOf(LETTER_E, pos + 1)) {
    if (bytes[pos + 1] === LETTER_I) return pos
  }
  return undefined
}

/** An `EI` at `from`, after optional whitespace: its offset. */
function markerAt(bytes: Uint8Array, from: number): number | undefined {
  let pos = from
  while (isWhitespace(bytes[pos])) pos += 1
  return bytes[pos] === LETTER_E && bytes[pos + 1] === LETTER_I ? pos : undefined
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
 * Just past the `EI` of an inline image whose data starts at `dataStart`, found as pdf.js finds it:
 *   - for JPEG and ASCII data, the first `EI` after the data's own end marker,
 *   - otherwise, or without that marker, the first `EI` that passes pdf.js's tests (`endsImage`).
 * pdf.js ignores a declared length (/L, from PDF 2.0) and the length unfiltered data must have,
 * but other readers go by them. When either points to a different `EI`, scanning stops, as it does
 * for an image without an end, rather than read its data as operators.
 */
function inlineImageEnd(bytes: Uint8Array, dataStart: number, entries: ReadonlyMap<string, string>): number {
  const dataEnd = endOfData(bytes, dataStart, firstName(entries.get('/F') ?? entries.get('/Filter')))
  const marker = dataEnd === undefined ? searchEndMarker(bytes, dataStart) : firstMarker(bytes, dataEnd)
  if (marker === undefined) throw new Error('A page contains an inline image whose end could not be found.')

  const declared = Number(entries.get('/L') ?? entries.get('/Length'))
  const length = Number.isInteger(declared) && declared >= 0 ? declared : unfilteredLength(entries)
  if (length !== undefined && markerAt(bytes, dataStart + length) !== marker) {
    throw new Error('A page contains an inline image whose length and data disagree, so readers could disagree about where it ends.')
  }
  // After an end marker, pdf.js takes the byte after `EI` too; that matters only when it is not a space.
  return dataEnd === undefined || bytes[marker + 2] === undefined || isWhitespace(bytes[marker + 2]) ? marker + 2 : marker + 3
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
