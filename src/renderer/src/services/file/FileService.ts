import type { FileReference, OpenFileOptions, SaveFileOptions } from '@shared/shell'

/** A dropped item that is not a readable file, such as a folder, and why. */
export interface SkippedItem {
  readonly name: string
  readonly reason: string
}

/** A drop's files, in the order dropped, and the items left out. */
export interface DroppedFiles {
  readonly files: readonly FileReference[]
  readonly skipped: readonly SkippedItem[]
}

export interface FileService {
  chooseFiles(options?: OpenFileOptions): Promise<readonly FileReference[]>
  /** Resolves to the chosen path, or undefined when the dialog is cancelled. */
  chooseSavePath(options: SaveFileOptions): Promise<string | undefined>
  describe(path: string): Promise<FileReference>
  describeDroppedFiles(files: ArrayLike<File>): Promise<DroppedFiles>
  read(path: string): Promise<Uint8Array>
  write(path: string, contents: Uint8Array): Promise<void>
}
