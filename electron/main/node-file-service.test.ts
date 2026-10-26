import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { chmod, lstat, mkdtemp, mkdir, readdir, readFile, rm, stat, symlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { NodeFileService, replaceFileSync } from './node-file-service'

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

  it('explains how to open a Numbers document saved as a package folder', async () => {
    const path = join(temporaryDirectory, 'Budget.numbers')
    await mkdir(path)

    await expect(service.describeFile(path)).rejects.toThrow('Change File Type > Single File')
  })

  it('leaves no temporary files behind after an atomic write', async () => {
    const path = join(temporaryDirectory, 'atomic.csv')
    await service.writeFile(path, Uint8Array.from([1, 2, 3]))

    expect(await readdir(temporaryDirectory)).toEqual(['atomic.csv'])
  })

  it('writes through a symlink to the file it points to, keeping the link', async () => {
    const real = join(temporaryDirectory, 'real.pdf')
    const link = join(temporaryDirectory, 'link.pdf')
    await writeFile(real, Uint8Array.of(1))
    await symlink(real, link)

    await service.writeFile(link, Uint8Array.of(2))

    expect((await lstat(link)).isSymbolicLink()).toBe(true)
    expect(new Uint8Array(await readFile(real))).toEqual(Uint8Array.of(2))
    expect((await readdir(temporaryDirectory)).sort()).toEqual(['link.pdf', 'real.pdf'])
  })

  it('keeps the original file\'s permissions', async () => {
    const path = join(temporaryDirectory, 'shared.docx')
    await writeFile(path, Uint8Array.of(1))
    await chmod(path, 0o640)

    await service.writeFile(path, Uint8Array.of(2))

    expect((await stat(path)).mode & 0o777).toBe(0o640)
  })

  it('removes the temporary file when the original cannot be replaced', async () => {
    await mkdir(join(temporaryDirectory, 'folder.pdf'))

    await expect(service.writeFile(join(temporaryDirectory, 'folder.pdf'), Uint8Array.of(1))).rejects.toThrow()
    expect(await readdir(temporaryDirectory)).toEqual(['folder.pdf'])
  })

  it('refuses to read or write unknown file types', async () => {
    const path = join(temporaryDirectory, 'secrets.txt')

    await expect(service.readFile(path)).rejects.toThrow('Unsupported file type: .txt')
    await expect(service.writeFile(path, Uint8Array.of(0))).rejects.toThrow('Unsupported file type: .txt')
    expect(await readdir(temporaryDirectory)).toEqual([])
  })

  it('replaces a settings file with complete JSON and no temporary file', async () => {
    const path = join(temporaryDirectory, 'session.json')
    replaceFileSync(path, JSON.stringify([{ paths: ['/old.pdf'] }]))
    replaceFileSync(path, JSON.stringify([{ paths: ['/new.pdf'] }]))

    expect(JSON.parse(await readFile(path, 'utf8'))).toEqual([{ paths: ['/new.pdf'] }])
    expect(await readdir(temporaryDirectory)).toEqual(['session.json'])
  })
})
