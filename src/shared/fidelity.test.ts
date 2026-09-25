import { describe, expect, it, vi } from 'vitest'
import { createImportReport } from './fidelity'

describe('createImportReport', () => {
  it('defaults an unspecified finding to degraded', () => {
    vi.spyOn(Date, 'now').mockReturnValue(123)

    expect(createImportReport('budget.xlsx', [{ construct: 'Unknown chart' }])).toEqual({
      id: 'budget.xlsx:123',
      sourceName: 'budget.xlsx',
      severity: 'degraded',
      findings: [{ construct: 'Unknown chart', severity: 'degraded' }],
    })
  })

  it('uses the highest severity and reports empty imports as lossless', () => {
    expect(createImportReport('clean.xlsx', []).severity).toBe('lossless')
    expect(createImportReport('macro.xlsm', [
      { construct: 'Merged cells' },
      { construct: 'VBA project', severity: 'dropped' },
    ]).severity).toBe('dropped')
  })
})
