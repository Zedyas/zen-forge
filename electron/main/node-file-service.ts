import { randomUUID } from 'node:crypto'
import { closeSync, fsyncSync, openSync, renameSync, rmSync, writeFileSync } from 'node:fs'
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

/**
 * Replaces a small file (settings, session) through a temporary file beside it, flushed to disk
 * before the rename, so a crash mid-write leaves either the old file or the new one, never half of
 * one. Synchronous, so writes land in the order they are made and a quit cannot cut one short.
 */
export function replaceFileSync(path: string, contents: string): void {
  const temporary = `${path}.tmp`
  try {
    const descriptor = openSync(temporary, 'w')
    try {
      writeFileSync(descriptor, contents)
      fsyncSync(descriptor)
    } finally {
      closeSync(descriptor)
    }
    renameSync(temporary, path)
  } catch (error) {
    rmSync(temporary, { force: true })
    throw error
  }
}

/** Main-process filesystem adapter used behind the FileService IPC contract. */
export class NodeFileService {
  async describeFile(path: string): Promise<FileReference> {
    const metadata = await stat(path)
    if (metadata.isDirectory() && extensionOf(path) === 'numbers') {
      // Older Numbers versions could save a document as a package: a folder that Finder shows as one file.
      throw new Error('This Numbers document is saved as a package. In Numbers, choose File > Advanced > Change File Type > Single File, then open it again.')
    }
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
