import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { mkdtemp, mkdir, readdir, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { NodeFileService } from './node-file-service'

describe('NodeFileService contract', () => {
  const service = new NodeFileService()
  let temporaryDirectory = ''

  beforeEach(async () => {
    temporaryDirectory = await mkdtemp(join(tmpdir(), 'zen-forge-file-service-'))
  })

  afterEach(async () => {
    await rm(temporaryDirectory, { recursive: true, force: true })
  })

  it('writes and reads bytes without changing them', async () => {
    const path = join(temporaryDirectory, 'Quarterly Report.xlsx')
    const contents = Uint8Array.from([0, 17, 128, 255])

    await service.writeFile(path, contents)

    await expect(service.readFile(path)).resolves.toEqual(contents)
    await expect(service.describeFile(path)).resolves.toMatchObject({
      path,
      name: 'Quarterly Report.xlsx',
      extension: 'xlsx',
      size: contents.length,
    })
  })

  it('rejects a directory passed as a file', async () => {
    const path = join(temporaryDirectory, 'not-a-file.pdf')
    await mkdir(path)

    await expect(service.describeFile(path)).rejects.toThrow('not a file')
  })

  it('leaves no temporary files behind after an atomic write', async () => {
    const path = join(temporaryDirectory, 'atomic.csv')
    await service.writeFile(path, Uint8Array.from([1, 2, 3]))

    expect(await readdir(temporaryDirectory)).toEqual(['atomic.csv'])
  })

  it('refuses to read or write unknown file types', async () => {
    const path = join(temporaryDirectory, 'secrets.txt')

    await expect(service.readFile(path)).rejects.toThrow('Unsupported file type: .txt')
    await expect(service.writeFile(path, Uint8Array.of(0))).rejects.toThrow('Unsupported file type: .txt')
    expect(await readdir(temporaryDirectory)).toEqual([])
  })
})
