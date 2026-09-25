import { resolve } from 'node:path'
import { defineConfig } from 'vitest/config'

// Mirrors the renderer aliases in electron.vite.config.ts so tests import modules the same way.
export default defineConfig({
  resolve: {
    alias: {
      '@renderer': resolve('src/renderer/src'),
      '@shared': resolve('src/shared'),
    },
  },
})
