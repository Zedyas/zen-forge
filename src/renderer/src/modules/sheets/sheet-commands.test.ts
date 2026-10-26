// @vitest-environment happy-dom
import { describe, expect, it } from 'vitest'
import { linkAddress } from './sheet-commands'

describe('linkAddress', () => {
  it.each([
    ['example.com', 'https://example.com'],
    ['https://example.com/a b', undefined],
    ['name@example.com', 'mailto:name@example.com'],
    ['#Totals!B1', '#Totals!B1'],
    ['localhost:3000', 'https://localhost:3000'],
    ['javascript:alert(1)', undefined],
    ['file:///etc/hosts', undefined],
    ['two words', undefined],
  ])('reads %s as %s', (input, link) => {
    expect(linkAddress(input)).toBe(link)
  })
})
