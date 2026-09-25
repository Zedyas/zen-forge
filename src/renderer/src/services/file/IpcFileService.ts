import type { FileReference, OpenFileOptions, SaveFileOptions, ShellBridge } from '@shared/shell'
import type { FileService } from './FileService'

/** Returns the isolated preload API; direct bridge access stays in this directory. */
export function getBridge(): ShellBridge | undefined {
  return window.desktop
}

export function requireBridge(action: string): ShellBridge {
  const bridge = getBridge()
  if (bridge === undefined) throw new Error(`${action} requires the desktop app.`)
  return bridge
}

/** Renderer filesystem adapter backed by typed IPC calls to the main process. */
export class IpcFileService implements FileService {
  chooseFiles(options: OpenFileOptions = {}): Promise<readonly FileReference[]> {
    return requireBridge('Opening files').showOpenDialog(options)
  }

  chooseSavePath(options: SaveFileOptions): Promise<string | undefined> {
    return requireBridge('Saving files').showSaveDialog(options)
  }

  describe(path: string): Promise<FileReference> {
    return requireBridge('Reading files').describeFile(path)
  }

  describeDroppedFiles(files: FileList): Promise<readonly FileReference[]> {
    const bridge = requireBridge('Dropped files')
    return Promise.all(Array.from(files).map(file => bridge.describeFile(bridge.pathForDroppedFile(file))))
  }

  read(path: string): Promise<Uint8Array> {
    return requireBridge('Reading files').readFile(path)
  }

  write(path: string, contents: Uint8Array): Promise<void> {
    return requireBridge('Writing files').writeFile(path, contents)
  }
}

export const fileService: FileService = new IpcFileService()
