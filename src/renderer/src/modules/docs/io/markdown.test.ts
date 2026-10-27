import { getSchema } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { documentExtensions } from '../schema'
import { markdownLosses, readMarkdown, writeMarkdown } from './markdown'

const source = `# Trip plan

Book the [train](https://example.com/trains) early.

- Passport
- Tickets
  - Return leg

#### Day trips

###### Optional

| Day | City |
| --- | --- |
| 1 | Kyoto |
| 2 | Nara |
`

describe('markdown', () => {
  it('opens an empty numbered item as an item with an empty paragraph to type in', () => {
    const { content } = readMarkdown('1. First\n2. \n3. Third\n')
    expect(() => getSchema(documentExtensions).nodeFromJSON(content).check()).not.toThrow()
    expect(content.content?.[0]?.content?.[1]).toEqual({ type: 'listItem', content: [{ type: 'paragraph' }] })
  })

  it('keeps headings, lists, tables and links through a save and reopen', () => {
    const opened = readMarkdown(source)
    const saved = writeMarkdown(opened.content)

    expect(opened.findings).toEqual([])
    expect(readMarkdown(saved).content).toEqual(opened.content)
    expect(opened.content.content?.map(node => [node.type, node.attrs?.['level']])).toEqual([
      ['heading', 1], ['paragraph', undefined], ['bulletList', undefined], ['heading', 4], ['heading', 6], ['table', undefined],
    ])
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

  it('keeps text that looks like Markdown as text, pipes in tables and fences inside code', () => {
    const lines = ['*not emphasis*', '# not a heading', '1. not a list', '- not a bullet', '[not](a-link)', '<b>not html</b>', '---']
    const paragraphs = lines.map(line => ({ type: 'paragraph', content: [{ type: 'text', text: line }] }))
    expect(readMarkdown(writeMarkdown({ type: 'doc', content: paragraphs })).content.content?.map(node => node.content?.[0]?.text)).toEqual(lines)

    for (const source of ['| Command | Meaning |\n| - | - |\n| `a \\| b` | pipe a \\| b |\n', 'Write:\n\n````md\n```js\nx\n```\n````\n']) {
      const opened = readMarkdown(source).content
      expect(readMarkdown(writeMarkdown(opened)).content).toEqual(opened)
    }
  })

  it('keeps each value of a merged table cell in its column', () => {
    const cell = (type: string, text: string, attrs: Record<string, unknown> = {}) => ({ type, attrs, content: [{ type: 'paragraph', content: [{ type: 'text', text }] }] })
    const saved = writeMarkdown({
      type: 'doc',
      content: [{
        type: 'table',
        content: [
          { type: 'tableRow', content: [cell('tableHeader', 'Region'), cell('tableHeader', 'Q1'), cell('tableHeader', 'Q2')] },
          { type: 'tableRow', content: [cell('tableCell', 'North', { colspan: 2 }), cell('tableCell', '42')] },
        ],
      }],
    })
    const row = readMarkdown(saved).content.content?.[0]?.content?.[1]?.content?.map(entry => entry.content?.[0]?.content?.[0]?.text ?? '')
    expect(row).toEqual(['North', '', '42'])
  })

  it('writes a table of contents as a list of links to the headings, and says so before saving over a file', () => {
    const heading = (level: number, text: string) => ({ type: 'heading', attrs: { level }, content: [{ type: 'text', text }] })
    const content = {
      type: 'doc',
      attrs: { headerFooter: { footer: [{ type: 'paragraph', content: [{ type: 'text', text: 'Draft' }] }] } },
      content: [{ type: 'tableOfContents', attrs: { levels: 2 } }, heading(1, 'Plan'), heading(2, 'Two words!'), heading(3, 'Too deep'), heading(1, 'Plan')],
    }
    expect(writeMarkdown(content)).toMatch(/^- \[Plan\]\(#plan\)\n\n? {2}- \[Two words!\]\(#two-words\)\n- \[Plan\]\(#plan-1\)\n\n# Plan\n/)
    expect(markdownLosses(content)).toEqual(['the header and footer are left out', 'the table of contents becomes a list of links that no longer updates'])
  })

  it('links a table of contents to the anchors GitHub gives every heading, the title too', () => {
    const heading = (level: number, text: string) => ({ type: 'heading', attrs: { level }, content: [{ type: 'text', text }] })
    const contents = { type: 'tableOfContents', attrs: { levels: 3 } }
    const title = { type: 'title', content: [{ type: 'text', text: 'Plan' }] }

    expect(writeMarkdown({ type: 'doc', content: [title, contents, heading(1, 'Plan'), heading(2, 'Plan'), heading(2, 'Plan-1')] }))
      .toMatch(/^# Plan\n\n- \[Plan\]\(#plan-1\)\n\n? {2}- \[Plan\]\(#plan-2\)\n {2}- \[Plan-1\]\(#plan-1-1\)\n/)
  })

  it('nests each table of contents entry under the last heading above its level', () => {
    const heading = (level: number, text: string) => ({ type: 'heading', attrs: { level }, content: [{ type: 'text', text }] })
    const content = { type: 'doc', content: [{ type: 'tableOfContents', attrs: { levels: 3 } }, heading(2, 'A'), heading(2, 'B'), heading(1, 'C'), heading(3, 'D'), heading(2, 'E')] }

    expect(writeMarkdown(content)).toMatch(/^- \[A\]\(#a\)\n- \[B\]\(#b\)\n- \[C\]\(#c\)\n\n? {2}- \[D\]\(#d\)\n {2}- \[E\]\(#e\)\n\n## A/)
  })

  it('keeps front matter exactly, and drops links that are not web, mail or relative', () => {
    const source = '---\ntitle: Trip\ntags: [a, b]\n---\n\n[click](javascript:alert(1)) and [notes](notes.md)\n'
    const opened = readMarkdown(source).content
    expect(writeMarkdown(opened)).toBe('---\ntitle: Trip\ntags: [a, b]\n---\n\nclick and [notes](notes.md)\n')
    expect(JSON.stringify(opened)).not.toContain('javascript:')
  })
})
