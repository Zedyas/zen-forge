import { afterEach, expect, it } from 'vitest'
import { formatCellValue } from './format'

const zone = process.env['TZ']
afterEach(() => {
  process.env['TZ'] = zone
})

it('shows a date typed as YYYY-MM-DD on that day west of Greenwich', () => {
  process.env['TZ'] = 'America/Chicago'
  expect(formatCellValue('2024-03-05', { numberFormat: 'date', decimalPlaces: 0 })).toBe(new Date(2024, 2, 5).toLocaleDateString())
})
