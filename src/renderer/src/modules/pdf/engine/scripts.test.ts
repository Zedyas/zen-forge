import { PDFDocument } from '@cantoo/pdf-lib'
import { describe, expect, it } from 'vitest'
import { savePdf } from './save'

async function withScripts(scripts: Readonly<Record<string, string>>): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  doc.addPage([400, 600])
  for (const [name, script] of Object.entries(scripts)) doc.addJavaScript(name, script)
  return doc.save()
}

describe('document scripts on combine', () => {
  it('keeps every file’s scripts once, renaming a different script with a taken name', async () => {
    const first = await withScripts({ shared: 'boilerplate()', check: 'first()' })
    const second = await withScripts({ shared: 'boilerplate()', check: 'second()', extra: 'extra()' })

    const output = await savePdf({
      sources: [first, second],
      pages: [{ source: 0, index: 0, rotation: 0, edits: [] }, { source: 1, index: 0, rotation: 0, edits: [] }],
    })

    const scripts = (await PDFDocument.load(output)).getDocumentJavaScripts()
    expect(scripts.map(({ name, script }) => [name, script])).toEqual([
      ['check', 'first()'],
      ['check (2)', 'second()'],
      ['extra', 'extra()'],
      ['shared', 'boilerplate()'],
    ])
  })
})
