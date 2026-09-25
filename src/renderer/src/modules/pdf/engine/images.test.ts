import { PDFDocument } from '@cantoo/pdf-lib'
import { describe, expect, it } from 'vitest'
import { imagesToPdf } from './images'
import { encodePng } from './test-support'

describe('imagesToPdf', () => {
  it('sizes each page from the image density rather than its pixel count', async () => {
    const output = await imagesToPdf([
      { bytes: encodePng(200, 100, [255, 0, 0], 144), type: 'png' },
      { bytes: encodePng(200, 100, [0, 255, 0]), type: 'png' },
    ])

    const sizes = (await PDFDocument.load(output)).getPages().map(page => page.getSize())
    expect(sizes[0].width).toBeCloseTo(100, 1)
    expect(sizes[0].height).toBeCloseTo(50, 1)
    expect(sizes[1]).toEqual({ width: 200, height: 100 })
  })
})
