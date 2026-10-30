/**
 * Excel number format codes (`#,##0.00`, `0%`, `yyyy-mm-dd`, `[Red]-0.0;0.0`, `0.00E+00`…), the common
 * subset: up to four sections (positive;negative;zero;text), named colours, quoted and escaped text,
 * currency and locale markers, dates and times, percent, thousands and scientific notation. Anything
 * else (conditions, elapsed time, fractions) is not supported: the cell shows as General, and the
 * code is still saved unchanged.
 */

export interface FormattedValue {
  readonly text: string
  /** `#rrggbb`, from a colour such as `[Red]`. */
  readonly color?: string
}

type Token =
  | { readonly kind: 'literal'; readonly text: string }
  | { readonly kind: 'digit'; readonly placeholder: '0' | '#' | '?' }
  | { readonly kind: 'point' }
  | { readonly kind: 'comma' }
  | { readonly kind: 'percent' }
  | { readonly kind: 'exponent'; readonly sign: '+' | '-' }
  | { readonly kind: 'text' }
  | { readonly kind: 'general' }
  | { readonly kind: 'date'; readonly part: string }

interface Section {
  readonly tokens: readonly Token[]
  readonly color?: string
  readonly date: boolean
}

/** Excel's eight named colours, as the darker shades the grid uses for text. */
const colors: Readonly<Record<string, string>> = {
  black: '#1d1f23', blue: '#2f5bd3', cyan: '#0e8fa3', green: '#1f8a4c',
  magenta: '#a8329b', red: '#c9352b', white: '#ffffff', yellow: '#b58a00',
}

const dayMilliseconds = 86_400_000
const serialEpoch = Date.UTC(1899, 11, 30)
/** Excel's limit on decimal places. */
const maximumDecimals = 30
/** Characters Excel shows as themselves without quotes. */
const plainLiterals = new Set([...'$-+/():!^&\'~{}<>= £€¥'])

/** Splits `positive;negative;zero;text` while ignoring separators inside quotes or brackets. */
export function splitSections(pattern: string): readonly string[] {
  const sections: string[] = []
  let current = ''
  let quoting = false
  let depth = 0
  let escaped = false
  for (const character of pattern) {
    if (escaped) {
      escaped = false
    } else if (character === '\\' && !quoting) {
      escaped = true
    } else if (character === '"') {
      quoting = !quoting
    } else if (!quoting && character === '[') {
      depth += 1
    } else if (!quoting && character === ']') {
      depth = Math.max(0, depth - 1)
    } else if (character === ';' && !quoting && depth === 0) {
      sections.push(current)
      current = ''
      continue
    }
    current += character
  }
  sections.push(current)
  return sections
}

/** The tokens of one section, or undefined when it uses something unsupported. */
function parseSection(source: string): Section | undefined {
  const tokens: Token[] = []
  let color: string | undefined
  let index = 0
  const literal = (text: string): void => {
    const last = tokens.at(-1)
    if (last?.kind === 'literal') tokens[tokens.length - 1] = { kind: 'literal', text: last.text + text }
    else tokens.push({ kind: 'literal', text })
  }
  while (index < source.length) {
    const rest = source.slice(index)
    const character = source[index] ?? ''
    const lower = rest.toLowerCase()
    if (character === '"') {
      const end = source.indexOf('"', index + 1)
      if (end < 0) return undefined
      literal(source.slice(index + 1, end))
      index = end + 1
    } else if (character === '\\') {
      literal(source[index + 1] ?? '')
      index += 2
    } else if (character === '_') {
      // Space as wide as the next character, which lines up positives with bracketed negatives.
      literal(' ')
      index += 2
    } else if (character === '*') {
      // Repeat the next character to fill the cell: nothing to fill in a grid that sizes text itself.
      index += 2
    } else if (character === '[') {
      const end = source.indexOf(']', index)
      if (end < 0) return undefined
      const inner = source.slice(index + 1, end)
      const named = colors[inner.toLowerCase()]
      if (named !== undefined) color = named
      else if (inner.startsWith('$')) literal(inner.slice(1).split('-')[0] ?? '')
      else return undefined
      index = end + 1
    } else if (lower.startsWith('general')) {
      tokens.push({ kind: 'general' })
      index += 7
    } else if (lower.startsWith('am/pm') || lower.startsWith('a/p')) {
      const length = lower.startsWith('am/pm') ? 5 : 3
      tokens.push({ kind: 'date', part: source.slice(index, index + length) })
      index += length
    } else if (/[ymdhs]/i.test(character)) {
      const run = rest.match(new RegExp(`^${character}+`, 'i'))?.[0] ?? character
      tokens.push({ kind: 'date', part: run.toLowerCase() })
      index += run.length
    } else if (character === '0' || character === '#' || character === '?') {
      tokens.push({ kind: 'digit', placeholder: character })
      index += 1
    } else if (character === '.') {
      tokens.push({ kind: 'point' })
      index += 1
    } else if (character === ',') {
      tokens.push({ kind: 'comma' })
      index += 1
    } else if (character === '%') {
      tokens.push({ kind: 'percent' })
      index += 1
    } else if ((character === 'E' || character === 'e') && (source[index + 1] === '+' || source[index + 1] === '-')) {
      tokens.push({ kind: 'exponent', sign: source[index + 1] === '+' ? '+' : '-' })
      index += 2
    } else if (character === '@') {
      tokens.push({ kind: 'text' })
      index += 1
    } else if (plainLiterals.has(character)) {
      literal(character)
      index += 1
    } else {
      return undefined
    }
  }
  const date = tokens.some(token => token.kind === 'date')
  // Fractions (# ?/?) and a date mixed with number placeholders are beyond this subset.
  const fraction = tokens.some((token, at) => token.kind === 'literal' && token.text.includes('/')
    && tokens[at - 1]?.kind === 'digit' && tokens[at + 1]?.kind === 'digit')
  if (fraction || (date && tokens.some(token => token.kind === 'digit' && token.placeholder !== '0'))) return undefined
  const pointAt = tokens.findIndex(token => token.kind === 'point')
  const exponentAt = tokens.findIndex(token => token.kind === 'exponent')
  const decimals = pointAt < 0 ? [] : tokens.slice(pointAt, exponentAt > pointAt ? exponentAt : undefined).filter(token => token.kind === 'digit')
  if (decimals.length > maximumDecimals) return undefined
  return { tokens, color, date }
}

function separators(): { readonly decimal: string; readonly group: string } {
  const parts = new Intl.NumberFormat(undefined).formatToParts(12_345.6)
  return {
    decimal: parts.find(part => part.type === 'decimal')?.value ?? '.',
    group: parts.find(part => part.type === 'group')?.value ?? ',',
  }
}

/** Digits into placeholders, right-aligned: the first placeholder takes any extra leading digits. */
function fillInteger(digits: string, placeholders: readonly ('0' | '#' | '?')[]): string[] {
  const filled = placeholders.map(() => '')
  const padded = digits === '0' ? '' : digits
  const offset = padded.length - placeholders.length
  placeholders.forEach((placeholder, position) => {
    const at = offset + position
    if (at >= 0) filled[position] = position === 0 ? padded.slice(0, at + 1) : padded[at] ?? ''
    else filled[position] = placeholder === '0' ? '0' : placeholder === '?' ? ' ' : ''
  })
  return filled
}

/**
 * `magnitude` (not negative) rounded half up to `places` decimals as Excel rounds: on its 15 significant
 * digits, not its binary value, so 2.675 (2.67499… in binary) shows as 2.68.
 */
function roundedDigits(magnitude: number, places: number): { readonly integer: string; readonly fraction: string } {
  const [mantissa = '0', exponent = '0'] = magnitude.toExponential(14).split('e')
  // The 15 digits are an integer scaled by 10^(exponent - 14); scale them to `places` decimals instead.
  const shift = Number(exponent) - 14 + places
  const digits = BigInt(mantissa.replace('.', ''))
  const divisor = 10n ** BigInt(Math.max(0, -shift))
  const scaled = shift >= 0 ? digits * 10n ** BigInt(shift) : (digits + divisor / 2n) / divisor
  const text = scaled.toString().padStart(places + 1, '0')
  return { integer: text.slice(0, text.length - places), fraction: text.slice(text.length - places) }
}

function formatNumber(value: number, section: Section, signed: boolean): string {
  const { tokens } = section
  const { decimal, group } = separators()
  const exponentAt = tokens.findIndex(token => token.kind === 'exponent')
  const pointAt = tokens.findIndex(token => token.kind === 'point')
  const numberEnd = exponentAt >= 0 ? exponentAt : tokens.length
  const placeholdersBetween = (start: number, end: number): ('0' | '#' | '?')[] => tokens.slice(start, end)
    .flatMap(token => token.kind === 'digit' ? [token.placeholder] : [])
  const integerEnd = pointAt >= 0 && pointAt < numberEnd ? pointAt : numberEnd
  const integerPlaceholders = placeholdersBetween(0, integerEnd)
  const fractionPlaceholders = pointAt >= 0 && pointAt < numberEnd ? placeholdersBetween(pointAt + 1, numberEnd) : []
  const exponentPlaceholders = exponentAt >= 0 ? placeholdersBetween(exponentAt + 1, tokens.length) : []
  const lastIntegerDigit = tokens.slice(0, integerEnd).findLastIndex(token => token.kind === 'digit')
  const firstIntegerDigit = tokens.findIndex(token => token.kind === 'digit')
  const grouping = tokens.slice(firstIntegerDigit, lastIntegerDigit).some(token => token.kind === 'comma')
  // Commas right after the number's last digit divide by a thousand each (`0.0,,"M"`).
  let scaleCommas = 0
  const lastDigit = tokens.slice(0, numberEnd).findLastIndex(token => token.kind === 'digit')
  for (let at = lastDigit + 1; tokens[at]?.kind === 'comma'; at += 1) scaleCommas += 1
  const percents = tokens.filter(token => token.kind === 'percent').length
  let magnitude = Math.abs(value) * 100 ** percents / 1000 ** scaleCommas

  let exponent = 0
  if (exponentAt >= 0 && magnitude !== 0) {
    exponent = Math.floor(Math.log10(magnitude)) - (Math.max(1, integerPlaceholders.length) - 1)
    magnitude /= 10 ** exponent
    if (Number(roundedDigits(magnitude, fractionPlaceholders.length).integer) >= 10 ** Math.max(1, integerPlaceholders.length)) {
      exponent += 1
      magnitude /= 10
    }
  }
  const { integer: integerDigits, fraction: fractionDigits } = roundedDigits(magnitude, fractionPlaceholders.length)
  // The sign is the value's own, as in Excel, whatever the shown digits round to.
  const negative = signed && value < 0

  const integerParts = fillInteger(integerDigits, integerPlaceholders)
  if (grouping) {
    const required = integerPlaceholders.filter(placeholder => placeholder === '0').length
    const whole = integerDigits === '0' && required === 0 ? '' : integerDigits.padStart(required, '0')
    integerParts.fill('')
    integerParts[0] = whole.replace(/\B(?=(\d{3})+(?!\d))/g, group)
  }
  // Optional fraction digits (# and ?) drop trailing zeros, from the right.
  const fractionParts = fractionPlaceholders.map((_, position) => fractionDigits[position] ?? '0')
  for (let position = fractionPlaceholders.length - 1; position >= 0; position -= 1) {
    const placeholder = fractionPlaceholders[position]
    if (placeholder === '0' || fractionParts[position] !== '0') break
    fractionParts[position] = placeholder === '?' ? ' ' : ''
  }
  const exponentText = String(Math.abs(exponent)).padStart(exponentPlaceholders.filter(placeholder => placeholder === '0').length, '0')

  let integerIndex = 0
  let fractionIndex = 0
  let text = ''
  tokens.forEach((token, at) => {
    switch (token.kind) {
      case 'literal':
        text += token.text
        return
      case 'percent':
        text += '%'
        return
      case 'point':
        if (at < numberEnd) text += decimal
        return
      case 'exponent':
        text += `E${exponent < 0 ? '-' : token.sign === '+' ? '+' : ''}${exponentText}`
        return
      case 'digit':
        if (at > exponentAt && exponentAt >= 0) return
        if (at < integerEnd) text += integerParts[integerIndex++] ?? ''
        else text += fractionParts[fractionIndex++] ?? ''
        return
      case 'general':
        text += String(Number(magnitude.toPrecision(10)))
        return
      case 'comma':
      case 'text':
      case 'date':
        return
    }
  })
  return negative ? `-${text}` : text
}

function datePart(part: string, date: Date, twelveHour: boolean, minuteContext: boolean): string {
  const month = date.getUTCMonth()
  const names = (style: 'long' | 'short', kind: 'month' | 'weekday'): string =>
    new Intl.DateTimeFormat(undefined, { [kind]: style, timeZone: 'UTC' }).format(date)
  const hours = date.getUTCHours()
  const shownHours = twelveHour ? (hours % 12 === 0 ? 12 : hours % 12) : hours
  const two = (value: number): string => String(value).padStart(2, '0')
  if (part.startsWith('y')) return part.length <= 2 ? two(date.getUTCFullYear() % 100) : String(date.getUTCFullYear())
  if (part.startsWith('m') && minuteContext && part.length <= 2) return part.length === 2 ? two(date.getUTCMinutes()) : String(date.getUTCMinutes())
  if (part === 'm') return String(month + 1)
  if (part === 'mm') return two(month + 1)
  if (part === 'mmm') return names('short', 'month')
  if (part === 'mmmmm') return names('long', 'month').charAt(0)
  if (part.startsWith('mmm')) return names('long', 'month')
  if (part === 'd') return String(date.getUTCDate())
  if (part === 'dd') return two(date.getUTCDate())
  if (part === 'ddd') return names('short', 'weekday')
  if (part.startsWith('d')) return names('long', 'weekday')
  if (part === 'h') return String(shownHours)
  if (part.startsWith('h')) return two(shownHours)
  if (part === 's') return String(date.getUTCSeconds())
  if (part.startsWith('s')) return two(date.getUTCSeconds())
  const pm = hours >= 12
  if (part.toLowerCase() === 'am/pm') return part === part.toLowerCase() ? (pm ? 'pm' : 'am') : (pm ? 'PM' : 'AM')
  return part === part.toLowerCase() ? (pm ? 'p' : 'a') : (pm ? 'P' : 'A')
}

/**
 * Excel rounds a date to the smallest unit its code shows: a fraction of a second (`ss.00`), a second
 * when it shows a time, and for a date alone only float noise, so 23:59:59.9 is still that day. The
 * date comes from the rounded value, so it rolls forward only once the time reaches midnight.
 */
function formatDate(serial: number, section: Section): string | undefined {
  const { tokens } = section
  const secondsPoint = tokens.findIndex((token, at) => {
    const before = tokens[at - 1]
    return token.kind === 'point' && before?.kind === 'date' && before.part.startsWith('s')
  })
  let decimals = 0
  while (secondsPoint >= 0 && decimals < 3 && tokens[secondsPoint + 1 + decimals]?.kind === 'digit') decimals += 1
  const showsTime = tokens.some(token => token.kind === 'date' && /^[hsa]/i.test(token.part))
  const unit = decimals > 0 ? 1000 / 10 ** decimals : showsTime ? 1000 : 1
  const date = new Date(serialEpoch + Math.round(serial * dayMilliseconds / unit) * unit)
  if (Number.isNaN(date.valueOf()) || serial < 0) return undefined
  const fraction = String(date.getUTCMilliseconds()).padStart(3, '0')
  const twelveHour = tokens.some(token => token.kind === 'date' && /^a/i.test(token.part))
  const timeParts = tokens.flatMap((token, at) => token.kind === 'date' ? [{ part: token.part, at }] : [])
  let text = ''
  tokens.forEach((token, at) => {
    if (token.kind === 'literal') text += token.text
    else if (token.kind === 'point') text += '.'
    else if (token.kind === 'comma') text += ','
    else if (token.kind === 'digit') text += at > secondsPoint && at <= secondsPoint + decimals ? fraction[at - secondsPoint - 1] : '0'
    else if (token.kind === 'date') {
      // `m` is minutes right after hours or right before seconds, as in Excel.
      const order = timeParts.findIndex(part => part.at === at)
      const minuteContext = timeParts[order - 1]?.part.startsWith('h') === true || timeParts[order + 1]?.part.startsWith('s') === true
      text += datePart(token.part, date, twelveHour, minuteContext)
    }
  })
  return text
}

function formatText(text: string, section: Section | undefined): string {
  if (section === undefined) return text
  return section.tokens.map(token => token.kind === 'text' ? text : token.kind === 'literal' ? token.text : '').join('')
}

const cache = new Map<string, readonly Section[] | undefined>()

function sectionsOf(code: string): readonly Section[] | undefined {
  if (!cache.has(code)) {
    const sections = splitSections(code).map(parseSection)
    cache.set(code, sections.length > 4 || sections.some(section => section === undefined) ? undefined : sections.filter(section => section !== undefined))
  }
  return cache.get(code)
}

/** Whether `formatWithCode` can show values in this code; otherwise the cell shows as General. */
export function supportsFormatCode(code: string): boolean {
  return code.trim() !== '' && sectionsOf(code) !== undefined
}

/** Whether the code shows numbers as dates or times. */
export function isDateCode(code: string): boolean {
  return sectionsOf(code)?.[0]?.date === true
}

/**
 * A value shown with an Excel format code, or undefined when the code is not supported. It never throws:
 * a code that fails on a value shows that value as General.
 */
export function formatWithCode(value: string | number | boolean | null, code: string): FormattedValue | undefined {
  try {
    return formatWithSections(value, code)
  } catch {
    return undefined
  }
}

function formatWithSections(value: string | number | boolean | null, code: string): FormattedValue | undefined {
  const sections = sectionsOf(code)
  if (sections === undefined || sections.length === 0) return undefined
  if (value === null) return { text: '' }
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE' }
  if (typeof value === 'string') {
    const textSection = sections[3] ?? sections.find(section => section.tokens.some(token => token.kind === 'text'))
    return { text: formatText(value, textSection), ...(textSection?.color === undefined ? {} : { color: textSection.color }) }
  }
  const numeric = sections.filter(section => !section.tokens.some(token => token.kind === 'text') || sections.length > 1)
  const section = value < 0 && numeric[1] !== undefined ? numeric[1]
    : value === 0 && numeric[2] !== undefined ? numeric[2]
      : numeric[0]
  if (section === undefined) return undefined
  // A negative's own section shows the number without the minus; the first section adds one.
  const signed = !(value < 0 && numeric[1] !== undefined)
  const text = section.date ? formatDate(value, section) : formatNumber(value, section, signed)
  return text === undefined ? undefined : { text, ...(section.color === undefined ? {} : { color: section.color }) }
}
