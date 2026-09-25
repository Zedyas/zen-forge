import { describe, expect, it } from 'vitest'
import { scanContent } from './content-stream'

/** Content as bytes, one byte per character, so tests can hold binary image data. */
function bytes(content: string): Uint8Array {
  return Uint8Array.from(content, character => character.charCodeAt(0))
}

describe('scanContent', () => {
  it('steps over strings, hex strings, dictionaries, arrays and comments that hold operator-like text', () => {
    const content = [
      '/Span <</ActualText (a >> b EMC) /Nested <</Kids [1 (q)]>>>> BDC',
      '(nested (paren\\) EMC) inside) Tj % EMC Do',
      '<45 4D 43> Tj [(Q) -12 (Q)] TJ EMC',
    ].join('\n')
    const instructions = scanContent(bytes(content))

    expect(instructions.map(instruction => instruction.operator)).toEqual(['BDC', 'Tj', 'Tj', 'TJ', 'EMC'])
    expect(instructions[0]?.operands).toEqual(['/Span', '<</ActualText (a >> b EMC) /Nested <</Kids [1 (q)]>>>>'])
    const [, show] = instructions
    expect(content.slice(show?.start, show?.end)).toBe('(nested (paren\\) EMC) inside) Tj')
  })

  it('decodes escapes in names', () => {
    expect(scanContent(bytes('/OC /Layer#201 BDC'))[0]?.operands).toEqual(['/OC', '/Layer 1'])
  })

  it.each([
    ['unfiltered data, measured from its size', 'BI /W 4 /H 1 /CS /G /BPC 8 ID EI (\nEI'],
    ['a declared length', 'BI /W 1 /H 1 /F /Fl /L 5 ID EI BT\nEI'],
    ['ASCII85 data, which ends at ~>', 'BI /W 1 /H 1 /F /A85 ID EI (9jqo^BlbD-~>\nEI'],
    ['compressed data, whose EI must be followed by text', 'BI /W 2 /H 2 /CS /RGB /BPC 8 /F /Fl ID x\x9cEI \x80\x81(\x82\nEI'],
  ])('steps over an inline image with %s', (_, image) => {
    const content = `q ${image} Q /OC /hidden BDC EMC`
    const instructions = scanContent(bytes(content))
    expect(instructions.map(instruction => instruction.operator)).toEqual(['q', 'BI', 'Q', 'BDC', 'EMC'])
    expect(content.slice(instructions[1]?.start, instructions[1]?.end)).toBe(image)
  })

  // Each image's data holds an `EI` written to end it early, so that ` EMC` would close a hidden block.
  it.each([
    ['a delimiter after the EI', 'BI /W 2 /H 2 /CS /G /BPC 8 /F /Fl ID \x80\x81EI/aaaaaaaaaaaaaaaaaa EMC \x80\x82\x83\nEI'],
    ['binary data after the EI, where a declared length points', 'BI /W 2 /H 2 /CS /G /BPC 8 /F /Fl /L 2 ID \x80\x81EI EMC \x80\x82\x83\x84\nEI'],
    ['a word that is no operator after the EI', 'BI /W 2 /H 2 /CS /G /BPC 8 /F /Fl ID \x80\x81EI aaaaaaaaaaaaaaaaaa EMC \x80\x82\x83\nEI'],
  ])('steps over an inline image whose data holds an EI with %s, as pdf.js does', (_, image) => {
    const content = `q ${image} Q BT (SURVIVOR) Tj ET`
    const instructions = scanContent(bytes(content))
    expect(instructions.map(instruction => instruction.operator)).toEqual(['q', 'BI', 'Q', 'BT', 'Tj', 'ET'])
    expect(content.slice(instructions[1]?.start, instructions[1]?.end)).toBe(image)
  })

  it('stops with an error at an inline image that has no end', () => {
    expect(() => scanContent(bytes('BI /W 1 /H 1 /F /Fl ID \x80\x81(\x82'))).toThrow()
  })
})
