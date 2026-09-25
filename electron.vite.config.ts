import tailwindcss from '@tailwindcss/vite'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'
import { resolve } from 'node:path'
import type { Plugin } from 'vite'

/**
 * The production page may load only Zendo's own files: no network, no eval. WebAssembly is allowed
 * for MuPDF, and inline styles for the grid and toasts. Development loads Vite's dev server, which
 * needs inline scripts, so the policy is added at build time only.
 */
const contentSecurityPolicy = [
  "default-src 'none'",
  "script-src 'self' 'wasm-unsafe-eval'",
  "worker-src 'self'",
  "style-src 'self' 'unsafe-inline'",
  "img-src 'self' data: blob:",
  "font-src 'self' data:",
  "connect-src 'self' data: blob:",
  "base-uri 'none'",
  "form-action 'none'",
  "object-src 'none'",
  "frame-src 'none'",
].join('; ')

function contentSecurityPolicyPlugin(): Plugin {
  return {
    name: 'zendo-content-security-policy',
    apply: 'build',
    transformIndexHtml: html => html.replace('<head>', `<head>\n    <meta http-equiv="Content-Security-Policy" content="${contentSecurityPolicy}" />`),
  }
}

export default defineConfig({
  main: {
    plugins: [externalizeDepsPlugin()],
    build: {
      rollupOptions: {
        input: resolve('electron/main/index.ts'),
      },
    },
  },
  preload: {
    plugins: [externalizeDepsPlugin()],
    build: {
      rollupOptions: {
        input: resolve('electron/preload/index.ts'),
      },
    },
  },
  renderer: {
    root: resolve('src/renderer'),
    resolve: {
      alias: {
        '@renderer': resolve('src/renderer/src'),
        '@shared': resolve('src/shared'),
      },
    },
    plugins: [react(), tailwindcss(), contentSecurityPolicyPlugin()],
  },
})
