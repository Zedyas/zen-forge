/// <reference types="vite/client" />

import type { ShellBridge } from '@shared/shell'

declare global {
  interface Window {
    readonly desktop?: ShellBridge
  }
}

export {}
