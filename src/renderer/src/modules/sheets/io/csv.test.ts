import { describe, expect, it } from 'vitest'
import { readCsv, writeCsv } from './csv'

/** windows-1252 is a single-byte encoding, so each character maps to its own code point. */
function encodeWindows1252(text: string): Uint8Array {
  return Uint8Array.from([...text], character => character.charCodeAt(0))
}

describe('csv', () => {
  it('reads a semicolon-delimited windows-1252 file with quoted newlines and typed numbers', () => {
    const bytes = encodeWindows1252(
      'Name;Amount;Note\nJosé;$1,234.50;"line one,\nline two"\n',
    )

    const { data, findings } = readCsv(bytes, 'Imported')
    const sheet = data.sheets[0]

    expect(sheet?.cells).toEqual([
      ['Name', 'Amount', 'Note'],
      ['José', 1234.5, 'line one,\nline two'],
    ])
    expect(sheet?.styles).toEqual([[1, 1, { numberFormat: 'currency', decimalPlaces: 2 }]])
    expect(findings.map(finding => finding.construct)).toEqual(['Non-Unicode text encoding'])
  })

  it('quotes separators, quotes and newlines and writes a UTF-8 BOM', () => {
    const bytes = writeCsv([
      ['plain', 'has,comma', 'has"quote'],
      ['has\nnewline', 'é', ''],
    ])

    expect([...bytes.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf])
    expect(new TextDecoder().decode(bytes.subarray(3))).toBe(
      'plain,"has,comma","has""quote"\r\n"has\nnewline",é,',
    )
  })
})
