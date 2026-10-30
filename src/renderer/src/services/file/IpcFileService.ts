import type { FileReference, OpenFileOptions, SaveFileOptions, ShellBridge } from '@shared/shell'
import type { DroppedFiles, FileService, SkippedItem } from './FileService'

/** Returns the isolated preload API; direct bridge access stays in this directory. */
export function getBridge(): ShellBridge | undefined {
  return window.desktop
}

export function requireBridge(action: string): ShellBridge {
  const bridge = getBridge()
  if (bridge === undefined) throw new Error(`${action} requires the desktop app.`)
  return bridge
}

/**
 * Electron prefixes an error thrown in the main process with "Error invoking remote method '<channel>':
 * Error: ". What follows is the message written for people, so that is what the renderer shows.
 */
export async function mainProcessMessage<T>(call: Promise<T>): Promise<T> {
  try {
    return await call
  } catch (error) {
    if (!(error instanceof Error)) throw error
    throw new Error(error.message.replace(/^Error invoking remote method '[^']*': (?:\w*Error: )?/, ''), { cause: error })
  }
}

/** Renderer filesystem adapter backed by typed IPC calls to the main process. */
export class IpcFileService implements FileService {
  chooseFiles(options: OpenFileOptions = {}): Promise<readonly FileReference[]> {
    return mainProcessMessage(requireBridge('Opening files').showOpenDialog(options))
  }

  chooseSavePath(options: SaveFileOptions): Promise<string | undefined> {
    return mainProcessMessage(requireBridge('Saving files').showSaveDialog(options))
  }

  describe(path: string): Promise<FileReference> {
    return mainProcessMessage(requireBridge('Reading files').describeFile(path))
  }

  /** One item that cannot be read, such as a folder, is skipped; the others still open. */
  async describeDroppedFiles(files: ArrayLike<File>): Promise<DroppedFiles> {
    const bridge = requireBridge('Dropped files')
    const items = Array.from(files)
    const results = await Promise.allSettled(items.map(file => mainProcessMessage(bridge.describeFile(bridge.pathForDroppedFile(file)))))
    const described: FileReference[] = []
    const skipped: SkippedItem[] = []
    results.forEach((result, index) => {
      if (result.status === 'fulfilled') described.push(result.value)
      else skipped.push({ name: items[index]?.name ?? '', reason: result.reason instanceof Error ? result.reason.message : 'It could not be read.' })
    })
    return { files: described, skipped }
  }

  read(path: string): Promise<Uint8Array> {
    return mainProcessMessage(requireBridge('Reading files').readFile(path))
  }

  write(path: string, contents: Uint8Array): Promise<void> {
    return mainProcessMessage(requireBridge('Writing files').writeFile(path, contents))
  }
}

export const fileService: FileService = new IpcFileService()
