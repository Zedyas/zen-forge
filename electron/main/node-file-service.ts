import { randomUUID } from 'node:crypto'
import { readFile, rename, rm, stat, writeFile } from 'node:fs/promises'
import { basename, extname } from 'node:path'
import { isReadableExtension, isWritableExtension } from '../../src/shared/applications'
import type { FileReference } from '../../src/shared/shell'

function extensionOf(path: string): string {
  return extname(path).replace(/^\./, '').toLowerCase()
}

function requireExtension(path: string, allowed: (extension: string) => boolean): void {
  const extension = extensionOf(path)
  if (!allowed(extension)) throw new Error(`Unsupported file type: .${extension || 'unknown'}`)
}

/** Main-process filesystem adapter used behind the FileService IPC contract. */
export class NodeFileService {
  async describeFile(path: string): Promise<FileReference> {
    const metadata = await stat(path)
    if (!metadata.isFile()) throw new Error('The selected path is not a file.')

    return {
      path,
      name: basename(path),
      extension: extensionOf(path),
      size: metadata.size,
      lastModified: metadata.mtimeMs,
    }
  }

  async readFile(path: string): Promise<Uint8Array> {
    requireExtension(path, isReadableExtension)
    return new Uint8Array(await readFile(path))
  }

  /** Writes through a temporary file so a crash mid-save cannot destroy the original. */
  async writeFile(path: string, contents: Uint8Array): Promise<void> {
    requireExtension(path, isWritableExtension)
    const temporary = `${path}.${randomUUID()}.tmp`
    try {
      await writeFile(temporary, contents)
      await rename(temporary, path)
    } catch (error) {
      await rm(temporary, { force: true }).catch(() => undefined)
      throw error
    }
  }
}
