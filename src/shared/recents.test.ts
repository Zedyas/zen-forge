import { describe, expect, it } from 'vitest'
import { recentMenuItems, recentMenuLimit } from './recents'

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
