import { afterEach, describe, expect, it, vi } from 'vitest'

const { app, handlers } = vi.hoisted(() => ({
  app: { isPackaged: true, on: () => undefined },
  handlers: new Map<string, (event: unknown) => unknown>(),
}))
vi.mock('electron', () => ({
  app,
  ipcMain: { handle: (channel: string, listener: (event: unknown) => unknown) => handlers.set(channel, listener) },
  net: {},
  protocol: {},
  session: {},
}))

import { appPageUrl, handle, rendererUrl } from './security'

const attacker = 'https://attacker.example'

function invokeFrom(url: string): unknown {
  handle('test:channel', () => 'answered')
  return handlers.get('test:channel')?.({ senderFrame: { parent: null, url } })
}

describe('the page Zen Suzu loads and answers', () => {
  afterEach(() => {
    vi.unstubAllEnvs()
    app.isPackaged = true
  })

  it('ignores the development server address in a packaged app', () => {
    // The fuses stop environment variables from changing the packaged app; this one must not either.
    vi.stubEnv('ELECTRON_RENDERER_URL', attacker)
    expect(rendererUrl()).toBe(appPageUrl)
    expect(() => invokeFrom(`${attacker}/`)).toThrow("not sent by Zen Suzu's own page")
    expect(invokeFrom(appPageUrl)).toBe('answered')
  })

  it('loads and answers the development server in development', () => {
    app.isPackaged = false
    vi.stubEnv('ELECTRON_RENDERER_URL', 'http://localhost:5173')
    expect(rendererUrl()).toBe('http://localhost:5173')
    expect(invokeFrom('http://localhost:5173/')).toBe('answered')
  })
})
