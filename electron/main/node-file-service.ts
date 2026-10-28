import { randomUUID } from 'node:crypto'
import { closeSync, fsyncSync, openSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { open, readFile, realpath, rename, rm, stat } from 'node:fs/promises'
import { basename, dirname, extname, join } from 'node:path'
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

  /**
   * Writes through a temporary file, flushed to disk before it replaces the original, so a crash
   * mid-save cannot destroy the original. A symlink is followed and the file it points to is
   * replaced, keeping the link; the new file keeps the original's permissions. The temporary file
   * is hidden and removed when the write fails.
   */
  async writeFile(path: string, contents: Uint8Array): Promise<void> {
    requireExtension(path, isWritableExtension)
    // A new file has nothing to resolve or keep.
    const target = await realpath(path).catch(() => path)
    const mode = (await stat(target).catch(() => undefined))?.mode
    const temporary = join(dirname(target), `.${basename(target)}.${randomUUID()}.tmp`)
    try {
      const handle = await open(temporary, 'wx')
      try {
        await handle.writeFile(contents)
        if (mode !== undefined) await handle.chmod(mode & 0o7777)
        await handle.sync()
      } finally {
        await handle.close()
      }
      await rename(temporary, target)
    } catch (error) {
      await rm(temporary, { force: true }).catch(() => undefined)
      throw error
    }
  }
}
