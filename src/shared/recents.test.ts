import { describe, expect, it } from 'vitest'
import { acceptRecents, recentMenuItems, recentMenuLimit, recentPathLimit } from './recents'

describe('recentMenuItems', () => {
  it('keeps the order it is given, lists each file once and stops at the limit', () => {
    const paths = Array.from({ length: 14 }, (_, index) => `/files/report-${index}.docx`)
    const items = recentMenuItems(['/files/report-3.docx', ...paths])
    expect(items).toHaveLength(recentMenuLimit)
    expect(items.map(item => item.path)).toEqual(['/files/report-3.docx', ...paths.filter(path => path !== '/files/report-3.docx')].slice(0, recentMenuLimit))
    expect(items[0]?.label).toBe('report-3.docx')
  })

  it('adds the folder to files that share a name, and the whole path when the folder is shared too', () => {
    const labels = recentMenuItems(['/work/q1/budget.xlsx', '/home/me/budget.xlsx', '/work/q1/notes.md', '/archive/q1/budget.xlsx']).map(item => item.label)
    expect(labels).toEqual(['budget.xlsx — /work/q1', 'budget.xlsx — me', 'notes.md', 'budget.xlsx — /archive/q1'])
  })
})

describe('acceptRecents', () => {
  const current = ['/files/a.docx']

  it('ignores a payload that is not a list, and entries that are not absolute paths to a file Zen Suzu opens', () => {
    expect(acceptRecents(current, null)).toBe(current)
    expect(acceptRecents(current, { length: 1, 0: '/files/b.docx' })).toBe(current)
    const next = acceptRecents(current, [null, 7, {}, '', 'b.docx', '../b.docx', '/files/notes.exe', '/files/no-extension', '/files/b\0.docx', '/files/b.docx', 'C:\\Files\\c.XLSX'])
    expect(next).toEqual(['/files/b.docx', 'C:\\Files\\c.XLSX'])
    expect(recentMenuItems(next).map(item => item.label)).toEqual(['b.docx', 'c.XLSX'])
  })

  it('keeps at most the limit and returns the current list itself when nothing changed', () => {
    const many = Array.from({ length: recentPathLimit + 5 }, (_, index) => `/files/report-${index}.pdf`)
    expect(acceptRecents(current, many)).toEqual(many.slice(0, recentPathLimit))
    expect(acceptRecents(many.slice(0, recentPathLimit), many)).toEqual(many.slice(0, recentPathLimit))
    expect(acceptRecents(current, [...current])).toBe(current)
  })

  it('accepts the empty list a window reports after Clear Menu, which leaves the menu without files', () => {
    const cleared = acceptRecents(current, [])
    expect(cleared).toEqual([])
    expect(recentMenuItems(cleared)).toEqual([])
  })
})
