import { zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { assertZipFitsInMemory } from './zip'

describe('assertZipFitsInMemory', () => {
  it('refuses a small zip that unpacks past the limit, and accepts an ordinary one', () => {
    // 1.1 GB of zeros would take real memory to build; claim the size in the directory instead.
    const bomb = zipSync({ 'a.xml': new Uint8Array(16) }, { level: 0 })
    const view = new DataView(bomb.buffer, bomb.byteOffset, bomb.byteLength)
    // The central directory entry (signature PK 1 2) holds the unpacked size at offset 24.
    const directory = bomb.findLastIndex((_, index) => index <= bomb.length - 4 && view.getUint32(index, true) === 0x02014b50)
    view.setUint32(directory + 24, 1.1 * 1024 ** 3, true)
    expect(() => assertZipFitsInMemory(bomb)).toThrow('too large to open safely')
    expect(() => assertZipFitsInMemory(zipSync({ 'a.xml': new Uint8Array(1000) }))).not.toThrow()
  })
})
