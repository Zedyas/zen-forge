// @vitest-environment happy-dom
import { getSchema, type JSONContent } from '@tiptap/core'
import { CommentRangeEnd, CommentRangeStart, CommentReference, Document, Header, Packer, Paragraph, TextRun } from 'docx'
import { describe, expect, it } from 'vitest'
import { documentExtensions } from '../schema'
import { defaultTheme, type DocumentTheme, type PageSetup } from '../theme'
import { writeDocx } from './docx-export'
import { readDocx } from './docx-read'

const schema = getSchema(documentExtensions)

/** The document as the editor holds it, with every attribute default filled in. */
function normalized(content: JSONContent): JSONContent {
  return schema.nodeFromJSON(content).toJSON()
}

function text(value: string, ...marks: Array<string | { readonly type: string; readonly attrs: Record<string, string> }>): JSONContent {
  return { type: 'text', text: value, ...(marks.length > 0 ? { marks: marks.map(mark => typeof mark === 'string' ? { type: mark } : mark) } : {}) }
}

function paragraph(...content: JSONContent[]): JSONContent {
  return { type: 'paragraph', content }
}

function item(...content: JSONContent[]): JSONContent {
  return { type: 'listItem', content }
}

function cell(type: string, value: string, attrs: Record<string, unknown> = {}): JSONContent {
  return { type, attrs, content: [paragraph(text(value))] }
}

/** A PNG header is all the writer reads: it takes the pixel size from it and embeds the bytes. */
const png = `data:image/png;base64,${btoa(String.fromCharCode(
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52,
  0, 0, 0, 40, 0, 0, 0, 20, 8, 6, 0, 0, 0, 0, 0, 0, 0,
))}`

const page: PageSetup = { width: 16_838, height: 11_906, marginTop: 1_080, marginRight: 1_440, marginBottom: 1_080, marginLeft: 1_800, pageNumbers: true }
const theme: DocumentTheme = { ...defaultTheme, normal: { ...defaultTheme.normal, fontFamily: 'Georgia', fontSize: 12, lineHeight: 1.5 } }

/** Everything the model holds, as Sumi would save it. */
const document: JSONContent = {
  type: 'doc',
  attrs: { page, theme },
  content: [
    { type: 'title', attrs: { textAlign: 'center' }, content: [text('Quarterly report')] },
    { type: 'heading', attrs: { level: 1 }, content: [text('Summary')] },
    { type: 'heading', attrs: { level: 2 }, content: [text('Regions')] },
    { type: 'heading', attrs: { level: 3 }, content: [text('North')] },
    {
      type: 'paragraph',
      attrs: { textAlign: 'justify', spaceBefore: 6, spaceAfter: 12, lineHeight: 2, indentLeft: 36, indentFirstLine: -18 },
      content: [
        text('Plain, '), text('bold', 'bold'), text(' '), text('italic', 'italic'), text(' '), text('underlined', 'underline'),
        text(' '), text('struck', 'strike'), text(' E=mc'), text('2', 'superscript'), text(' H'), text('2', 'subscript'), text('O '),
        text('x = 1', 'code'), text(' '), text('Verdana', { type: 'textStyle', attrs: { fontFamily: 'Verdana', fontSize: '14pt', color: '#c00000' } }),
        text(' '), text('highlighted', { type: 'textStyle', attrs: { backgroundColor: '#ffff00' } }),
        text(' '), text('shaded', { type: 'textStyle', attrs: { backgroundColor: '#dbe8fb' } }),
        text(' and a '), text('link', { type: 'link', attrs: { href: 'https://example.com/' } }), text('.'),
      ],
    },
    paragraph(text('Name\tValue'), { type: 'hardBreak' }, text('Second line')),
    {
      type: 'bulletList',
      content: [
        item(paragraph(text('First point')), paragraph(text('More about it'))),
        item(paragraph(text('Second point')), { type: 'orderedList', attrs: { start: 3 }, content: [item(paragraph(text('Step three')))] }),
      ],
    },
    { type: 'orderedList', content: [item(paragraph(text('One'))), item(paragraph(text('Two')))] },
    {
      type: 'taskList',
      content: [
        { type: 'taskItem', attrs: { checked: true }, content: [paragraph(text('Done'))] },
        { type: 'taskItem', attrs: { checked: false }, content: [paragraph(text('Open'))] },
      ],
    },
    { type: 'blockquote', content: [paragraph(text('A quotation')), paragraph(text('in two paragraphs'))] },
    { type: 'codeBlock', content: [text('const total = 42\n\nreturn total')] },
    { type: 'horizontalRule' },
    {
      type: 'table',
      content: [
        { type: 'tableRow', content: [cell('tableHeader', 'Region', { colwidth: [160] }), cell('tableHeader', 'Revenue', { colwidth: [120] }), cell('tableHeader', 'Share', { colwidth: [100] })] },
        { type: 'tableRow', content: [cell('tableCell', 'North', { rowspan: 2, colwidth: [160] }), cell('tableCell', '42', { colspan: 2, colwidth: [120, 100], backgroundColor: '#fde68a' })] },
        { type: 'tableRow', content: [cell('tableCell', '38', { colwidth: [120] }), cell('tableCell', '48%', { colwidth: [100] })] },
      ],
    },
    paragraph(text('A picture '), { type: 'image', attrs: { src: png, width: 40, height: 20 } }),
    { type: 'pageBreak' },
    paragraph(text('After the break')),
  ],
}

describe('docx', () => {
  it('opens a file Sumi saved exactly as it was saved', async () => {
    const { bytes, skippedImages } = await writeDocx(document)
    const read = readDocx(bytes)

    expect(skippedImages).toBe(0)
    expect(read.findings).toEqual([])
    expect(normalized(read.content)).toEqual(normalized(document))
  })

  it('lists comments and header text as lost, with counts and what happens to them', async () => {
    const source = new Document({
      comments: { children: [0, 1].map(id => ({ id, author: 'Reviewer', date: new Date(0), children: [new Paragraph('Check this')] })) },
      sections: [{
        headers: { default: new Header({ children: [new Paragraph('Confidential')] }) },
        children: [0, 1].map(id => new Paragraph({
          children: [new CommentRangeStart(id), new TextRun('Reviewed text'), new CommentRangeEnd(id), new TextRun({ children: [new CommentReference(id)] })],
        })),
      }],
    })
    const { findings } = readDocx(new Uint8Array(await Packer.toArrayBuffer(source)))

    expect(findings).toContainEqual({ construct: 'Comments', severity: 'dropped', location: '2 comments', suggestedAlternative: 'Removed when saved; the original file keeps them.' })
    expect(findings).toContainEqual(expect.objectContaining({ construct: 'Header and footer text', severity: 'dropped' }))
  })
})
