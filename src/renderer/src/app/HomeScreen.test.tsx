// @vitest-environment happy-dom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import type { RecentDocument } from '../services/index/document-index'

const { openRecent, openFiles } = vi.hoisted(() => ({
  openRecent: vi.fn<(path: string) => Promise<void>>(async () => undefined),
  openFiles: vi.fn(async () => undefined),
}))
// Stored when the file was last opened here; another app has changed it since.
const stored: RecentDocument = { path: '/files/report.docx', name: 'report.docx', extension: 'docx', size: 1, lastModified: 0, openedAt: 0 }
vi.mock('./document-actions', () => ({ openRecent, openFiles, chooseAndOpenFiles: async () => undefined }))
vi.mock('../services/index/document-index', () => ({ listRecents: async () => [stored], watchRecents: () => () => undefined, forgetRecent: async () => undefined }))
vi.mock('../services/platform/client', () => ({ platformClient: { getAppearance: async () => 'system' } }))
vi.mock('../modules/pdf/workflows', () => ({ combinePdfs: async () => undefined, createPdfFromImages: async () => undefined }))
vi.mock('./updates', () => ({ UpdateOffer: () => null }))

import { HomeScreen } from './HomeScreen'

Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })

describe('a recent file on Home', () => {
  it('opens through the file on disk, not the size and date stored when it was last opened', async () => {
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)
    await act(async () => root.render(<HomeScreen />))
    act(() => host.querySelector<HTMLButtonElement>('.recent-card')?.click())
    expect(openRecent).toHaveBeenCalledWith(stored.path)
    expect(openFiles).not.toHaveBeenCalled()
    act(() => root.unmount())
    host.remove()
  })
})
