// @vitest-environment happy-dom
import { generateJSON, getSchema, type JSONContent } from '@tiptap/core'
import { Document, Header, Packer, Paragraph, TextRun, CommentRangeEnd, CommentRangeStart, CommentReference } from 'docx'
import { describe, expect, it } from 'vitest'
import { letterPage } from '../page'
import { documentExtensions } from '../schema'
import { writeDocx } from './docx-export'
import { readDocx } from './docx-import'

const schema = getSchema(documentExtensions)

/** The document as the editor holds it, with every attribute default filled in. */
function normalized(content: JSONContent): JSONContent {
  return schema.nodeFromJSON(content).toJSON()
}

function text(value: string, ...marks: string[]): JSONContent {
  return { type: 'text', text: value, ...(marks.length > 0 ? { marks: marks.map(type => ({ type })) } : {}) }
}

function paragraph(...content: JSONContent[]): JSONContent {
  return { type: 'paragraph', content }
}

function list(type: string, ...items: string[]): JSONContent {
  return { type, content: items.map(item => ({ type: 'listItem', content: [paragraph(text(item))] })) }
}

function row(...cells: string[]): JSONContent {
  return { type: 'tableRow', content: cells.map(cell => ({ type: 'tableCell', content: [paragraph(text(cell))] })) }
}

/** A PNG header is all the writer reads: it takes the pixel size from it and embeds the bytes. */
const png = `data:image/png;base64,${btoa(String.fromCharCode(
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52,
  0, 0, 0, 40, 0, 0, 0, 20, 8, 6, 0, 0, 0, 0, 0, 0, 0,
))}`

const document: JSONContent = {
  type: 'doc',
  content: [
    { type: 'heading', attrs: { level: 1 }, content: [text('Quarterly report')] },
    { type: 'heading', attrs: { level: 2 }, content: [text('Summary')] },
    {
      type: 'paragraph',
      attrs: { textAlign: 'center' },
      content: [
        text('Plain, '), text('bold', 'bold'), text(', '), text('italic', 'italic'), text(', '), text('underlined', 'underline'),
        text(' and '), text('struck', 'strike'), text(' with a '),
        { type: 'text', text: 'link', marks: [{ type: 'link', attrs: { href: 'https://example.com/' } }] }, text('.'),
      ],
    },
    list('bulletList', 'First point', 'Second point'),
    list('orderedList', 'Step one', 'Step two'),
    { type: 'table', content: [row('Region', 'Revenue'), row('North', '42')] },
    { type: 'pageBreak' },
    { type: 'paragraph', attrs: { textAlign: 'right' }, content: [text('After the break')] },
    { type: 'blockquote', content: [paragraph(text('A quotation'))] },
    { type: 'codeBlock', content: [text('const total = 42\nreturn total')] },
    {
      type: 'taskList',
      content: [
        { type: 'taskItem', attrs: { checked: true }, content: [paragraph(text('Done'))] },
        { type: 'taskItem', attrs: { checked: false }, content: [paragraph(text('Open'))] },
      ],
    },
    { type: 'horizontalRule' },
    paragraph(text('Code like '), text('x = 1', 'code'), text(' and a picture '), { type: 'image', attrs: { src: png, width: 40, height: 20 } }),
  ],
}

describe('docx', () => {
  it('keeps structure, formatting, alignment and page breaks through a save and reopen', async () => {
    const { bytes, skippedImages } = await writeDocx(document, letterPage)
    const imported = await readDocx(bytes, letterPage)

    expect(skippedImages).toBe(0)
    // A file Sumi wrote holds nothing Sumi cannot show, so reopening it reports no loss.
    expect(imported.findings).toEqual([])
    expect(normalized(generateJSON(imported.html, documentExtensions))).toEqual(normalized(document))
  })

  it('reports comments and headers as lost, and colours as approximated', async () => {
    const source = new Document({
      comments: { children: [{ id: 0, author: 'Reviewer', date: new Date(0), children: [new Paragraph('Check this')] }] },
      sections: [{
        headers: { default: new Header({ children: [new Paragraph('Confidential')] }) },
        children: [
          new Paragraph({ children: [new CommentRangeStart(0), new TextRun('Reviewed text'), new CommentRangeEnd(0), new TextRun({ children: [new CommentReference(0)] })] }),
          new Paragraph({ children: [new TextRun({ text: 'Red text', color: 'C00000' })] }),
        ],
      }],
    })
    const { findings } = await readDocx(new Uint8Array(await Packer.toArrayBuffer(source)), letterPage)
    const severity = (construct: string): string | undefined => findings.find(finding => finding.construct === construct)?.severity

    expect(severity('Comments')).toBe('dropped')
    expect(severity('Headers and footers')).toBe('dropped')
    expect(severity('Text colours')).toBe('degraded')
  })
})
