import { describe, expect, it } from 'vitest'
import { readDelimited, writeDelimited } from './csv'

/** windows-1252 is a single-byte encoding, so each character maps to its own code point. */
function encodeWindows1252(text: string): Uint8Array {
  return Uint8Array.from([...text], character => character.charCodeAt(0))
}

describe('csv', () => {
  it('reads a semicolon-delimited windows-1252 file with quoted newlines and typed numbers', () => {
    const bytes = encodeWindows1252(
      'Name;Amount;Note\nJosé;$1,234.50;"line one,\nline two"\n',
    )

    const { data, findings } = readDelimited(bytes, 'Imported', 'csv')
    const sheet = data.sheets[0]

    expect(sheet?.cells).toEqual([
      ['Name', 'Amount', 'Note'],
      ['José', 1234.5, 'line one,\nline two'],
    ])
    expect(sheet?.styles).toEqual([[1, 1, { numberFormat: 'currency', decimalPlaces: 2 }]])
    expect(findings.map(finding => finding.construct)).toEqual(['Non-Unicode text encoding'])
  })

  it('quotes separators, quotes and newlines and writes a UTF-8 BOM', () => {
    const bytes = writeDelimited([
      ['plain', 'has,comma', 'has"quote'],
      ['has\nnewline', 'é', ''],
    ], 'csv')

    expect([...bytes.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf])
    expect(new TextDecoder().decode(bytes.subarray(3))).toBe(
      'plain,"has,comma","has""quote"\r\n"has\nnewline",é,',
    )
  })
})

describe('tsv', () => {
  it('round-trips tabs, quotes and newlines inside fields and leaves commas unquoted', () => {
    const rows = [
      ['Item', 'Note', 'Cost'],
      ['Desk, oak', 'left\tright', '120'],
      ['Lamp', 'say "hi"\non two lines', '35.5'],
    ]

    const bytes = writeDelimited(rows, 'tsv')
    expect(new TextDecoder().decode(bytes.subarray(3))).toBe(
      'Item\tNote\tCost\r\nDesk, oak\t"left\tright"\t120\r\nLamp\t"say ""hi""\non two lines"\t35.5',
    )

    const { data, findings } = readDelimited(bytes, 'Imported', 'tsv')
    expect(data.sheets[0]?.cells).toEqual([
      ['Item', 'Note', 'Cost'],
      ['Desk, oak', 'left\tright', 120],
      ['Lamp', 'say "hi"\non two lines', 35.5],
    ])
    expect(findings).toEqual([])
  })
})
