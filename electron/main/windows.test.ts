import { describe, expect, it } from 'vitest'
import { targetFor, windowWith, type WindowFacts } from './windows'

const report = '/files/report.docx'

function windowWithTabs(paths: readonly string[], crashed = false): WindowFacts {
  return { window: { webContents: { isCrashed: () => crashed } }, state: { session: { paths } } }
}

describe('the window a file goes to', () => {
  it('is the window that already has the file, not the current one', () => {
    const holder = windowWithTabs([report])
    const current = windowWithTabs([])
    expect(targetFor([holder, current], current, report)).toBe(holder)
  })

  it('is a window whose page still runs when the window that had the file crashed', () => {
    const crashed = windowWithTabs([report], true)
    const other = windowWithTabs([])
    expect(targetFor([crashed, other], crashed, report)).toBe(other)
    expect(targetFor([crashed], crashed, report)).toBeUndefined()
  })

  it('is never the asking window itself, so a file it has open is opened there as usual', () => {
    const sender = windowWithTabs([report])
    expect(windowWith([sender], report, sender.window.webContents)).toBeUndefined()
  })
})
