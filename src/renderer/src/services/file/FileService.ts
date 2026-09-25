import type { FileReference, OpenFileOptions, SaveFileOptions } from '@shared/shell'

export interface FileService {
  chooseFiles(options?: OpenFileOptions): Promise<readonly FileReference[]>
  /** Resolves to the chosen path, or undefined when the dialog is cancelled. */
  chooseSavePath(options: SaveFileOptions): Promise<string | undefined>
  describe(path: string): Promise<FileReference>
  describeDroppedFiles(files: FileList): Promise<readonly FileReference[]>
  read(path: string): Promise<Uint8Array>
  write(path: string, contents: Uint8Array): Promise<void>
}
