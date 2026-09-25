import { resolve } from 'node:path'
import { defineConfig } from 'vitest/config'

// Mirrors the renderer aliases in electron.vite.config.ts so tests import modules the same way.
export default defineConfig({
  resolve: {
    alias: {
      '@renderer': resolve('src/renderer/src'),
      '@shared': resolve('src/shared'),
      // The app bundles mammoth's browser build, which reads an ArrayBuffer; its Node build reads
      // only Node Buffers, so tests load mammoth's own prebuilt browser bundle instead.
      mammoth: resolve('node_modules/mammoth/mammoth.browser.js'),
    },
  },
})
