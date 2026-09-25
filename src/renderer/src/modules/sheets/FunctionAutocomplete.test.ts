import { describe, expect, it } from 'vitest'
import { functionPrefix } from './FunctionAutocomplete'

describe('functionPrefix', () => {
  it('finds the function name being typed, including the first one after =', () => {
    expect(functionPrefix('=SU')).toBe('SU')
    expect(functionPrefix('=1+SU')).toBe('SU')
    expect(functionPrefix('=A1 & LE')).toBe('LE')
    expect(functionPrefix('=SUM(A1, AV')).toBe('AV')
    expect(functionPrefix('SU')).toBeUndefined()
    expect(functionPrefix('="Paid on ti')).toBeUndefined()
    expect(functionPrefix('="a" & LE')).toBe('LE')
  })
})
