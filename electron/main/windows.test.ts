import { describe, expect, it } from 'vitest'
import { closeWaitsForPage, historyItems, targetFor, windowWith, type WindowFacts } from './windows'

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

describe('closing a window with unsaved documents', () => {
  it('waits for the page to ask, unless the page crashed and cannot answer', () => {
    const closing = (crashed: boolean) => ({ window: { webContents: { isCrashed: () => crashed } }, state: { hasUnsaved: true }, closeApproved: false })
    expect(closeWaitsForPage(closing(false))).toBe(true)
    expect(closeWaitsForPage(closing(true))).toBe(false)
  })
})

describe('Undo and Redo in the Edit menu', () => {
  it("undo typing in Home's text fields when no document has a history", () => {
    expect(historyItems([{ id: 'edit.find', label: 'Find' }])).toEqual([{ role: 'undo' }, { role: 'redo' }])
  })

  it("undo the active document's changes when it has them", () => {
    const undo = { id: 'edit.undo', label: 'Undo' }
    const redo = { id: 'edit.redo', label: 'Redo' }
    expect(historyItems([undo, redo, { id: 'edit.find', label: 'Find' }])).toEqual([undo, redo])
  })
})
