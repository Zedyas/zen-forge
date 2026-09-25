import { describe, expect, it } from 'vitest'
import { readMarkdown, writeMarkdown } from './markdown'

const source = `# Trip plan

Book the [train](https://example.com/trains) early.

- Passport
- Tickets
  - Return leg

| Day | City |
| --- | --- |
| 1 | Kyoto |
| 2 | Nara |
`

describe('markdown', () => {
  it('keeps headings, lists, tables and links through a save and reopen', () => {
    const opened = readMarkdown(source)
    const saved = writeMarkdown(opened.content)

    expect(opened.findings).toEqual([])
    expect(readMarkdown(saved).content).toEqual(opened.content)
    expect(opened.content.content?.map(node => node.type)).toEqual(['heading', 'paragraph', 'bulletList', 'table'])
    expect(saved).toContain('[train](https://example.com/trains)')
    expect(saved).toContain('| 2 ')
  })

  it('writes only what Markdown holds, without adding rows to tables', () => {
    const cell = (text: string) => ({ type: 'tableCell', content: [{ type: 'paragraph', content: [{ type: 'text', text }] }] })
    const saved = writeMarkdown({
      type: 'doc',
      content: [
        { type: 'paragraph', content: [{ type: 'text', text: 'Signed', marks: [{ type: 'underline' }, { type: 'textStyle', attrs: { color: '#c00000' } }] }] },
        { type: 'pageBreak' },
        { type: 'table', content: [{ type: 'tableRow', content: [cell('Day'), cell('City')] }, { type: 'tableRow', content: [cell('1'), cell('Kyoto')] }] },
      ],
    })
    expect(saved).toMatch(/^Signed\n+\| Day /)
    expect(readMarkdown(saved).content.content?.[1]?.content).toHaveLength(2)
  })
})
