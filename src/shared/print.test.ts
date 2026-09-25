import { describe, expect, it } from 'vitest'
import { paperForRegion } from './print'

describe('paperForRegion', () => {
  it('uses Letter in the Americas regions that use it and A4 everywhere else', () => {
    expect(['US', 'ca', 'MX'].map(region => paperForRegion(region).name)).toEqual(['Letter', 'Letter', 'Letter'])
    expect(['GB', 'DE', 'JP', ''].map(region => paperForRegion(region).name)).toEqual(['A4', 'A4', 'A4', 'A4'])
  })
})
