import { describe, expect, it } from 'vitest'
import { localImagePath } from './local-images'

describe('Markdown image paths', () => {
  it('reads only images inside the Markdown file’s own folder', () => {
    const folder = '/Users/me/Downloads/'
    expect(localImagePath(folder, 'images/chart.png')).toBe('/Users/me/Downloads/images/chart.png')
    expect(localImagePath(folder, 'images/../chart%201.png')).toBe('/Users/me/Downloads/chart 1.png')
    expect(localImagePath(folder, 'growth-50%.png')).toBe('/Users/me/Downloads/growth-50%.png')
    for (const outside of ['../../Desktop/passport.jpg', '/Users/victim/id.png', '~/secret.png', 'file:///etc/x.png', 'a/../../b.png']) {
      expect(localImagePath(folder, outside)).toBeUndefined()
    }
  })
})
