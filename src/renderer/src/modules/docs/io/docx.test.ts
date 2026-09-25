// @vitest-environment happy-dom
import { getSchema, type JSONContent } from '@tiptap/core'
import { CommentRangeEnd, CommentRangeStart, CommentReference, Document, Header, Packer, Paragraph, TextRun } from 'docx'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
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
  attrs: { page, theme, properties: { title: 'Quarterly report', subject: '', creator: 'Ana', keywords: 'sales', description: '' } },
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
        text(' and a '), text('link', { type: 'link', attrs: { href: 'https://example.com/' } }), text(' in '),
        text('red', { type: 'link', attrs: { href: 'https://example.com/#pricing' } }, { type: 'textStyle', attrs: { color: '#c00000' } }), text('.'),
      ],
    },
    paragraph(text('Name\tValue'), { type: 'hardBreak' }, text('Second line')),
    {
      type: 'bulletList',
      content: [
        item({ type: 'paragraph', attrs: { spaceAfter: 20, lineHeight: 2 }, content: [text('First point')] }, paragraph(text('More about it'))),
        item(
          paragraph(text('Second point')),
          { type: 'orderedList', attrs: { start: 3 }, content: [item(paragraph(text('Step three')))] },
          paragraph(text('Back in the second point')),
          { type: 'codeBlock', content: [text('npm test')] },
        ),
      ],
    },
    { type: 'orderedList', content: [item(paragraph(text('One'))), item(paragraph(text('Two')))] },
    { type: 'orderedList', content: [item(paragraph(text('A separate list')))] },
    {
      type: 'taskList',
      content: [
        { type: 'taskItem', attrs: { checked: true }, content: [paragraph(text('Done'))] },
        {
          type: 'taskItem',
          attrs: { checked: false },
          content: [paragraph(text('Open')), { type: 'taskList', content: [{ type: 'taskItem', attrs: { checked: true }, content: [paragraph(text('Nested step'))] }] }],
        },
      ],
    },
    {
      type: 'blockquote',
      content: [paragraph(text('A quotation')), paragraph(text('in two paragraphs')), { type: 'bulletList', content: [item(paragraph(text('with a list')))] }],
    },
    { type: 'codeBlock', content: [text('const total = 42\n\nreturn total')] },
    { type: 'horizontalRule' },
    {
      type: 'table',
      content: [
        { type: 'tableRow', content: [cell('tableHeader', 'Region', { colwidth: [160] }), cell('tableHeader', 'Revenue', { colwidth: [120] }), cell('tableHeader', 'Share', { colwidth: [100] })] },
        { type: 'tableRow', content: [cell('tableCell', 'North', { rowspan: 2, colwidth: [160] }), cell('tableCell', '42', { colspan: 2, colwidth: [120, 100], backgroundColor: '#fde68a' })] },
        {
          type: 'tableRow',
          content: [
            cell('tableCell', '38', { colwidth: [120] }),
            { type: 'tableCell', attrs: { colwidth: [100] }, content: [{ type: 'paragraph', attrs: { lineHeight: 2, indentLeft: 18 }, content: [text('48%')] }] },
          ],
        },
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

    expect(findings).toContainEqual({ construct: 'Comments', severity: 'dropped', location: '2 comments', suggestedAlternative: 'Not shown, and not kept when saved.' })
    expect(findings).toContainEqual(expect.objectContaining({ construct: 'Header and footer text', severity: 'dropped' }))
  })

  it('writes a file that opens when the text holds control characters', async () => {
    const { bytes } = await writeDocx({ type: 'doc', content: [paragraph(text('Page one\fPage two\u000bthree\u0007'))] })
    const xml = strFromU8(unzipSync(bytes)['word/document.xml'] ?? new Uint8Array())
    // eslint-disable-next-line no-control-regex
    expect(xml).not.toMatch(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/)
    const reopened = readDocx(bytes).content
    expect(reopened.content?.map(node => node.type)).toEqual(['paragraph', 'pageBreak', 'paragraph'])
    expect(JSON.stringify(reopened)).toContain('three')
  })
})

/** A hand-built package: one document part, plus any other parts. */
function packageOf(body: string, parts: Record<string, string> = {}): Uint8Array {
  const namespaces = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
  return zipSync({
    'word/document.xml': strToU8(`<?xml version="1.0"?><w:document ${namespaces}><w:body>${body}</w:body></w:document>`),
    ...Object.fromEntries(Object.entries(parts).map(([name, xml]) => [name, strToU8(xml.replace('NS', namespaces))])),
  })
}

const run = (value: string): string => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`
const para = (inner: string, pPr = ''): string => `<w:p>${pPr === '' ? '' : `<w:pPr>${pPr}</w:pPr>`}${inner}</w:p>`
const tc = (value: string): string => `<w:tc>${para(run(value))}</w:tc>`
const listItem = (value: string, numId: string): string => para(run(value), `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`)
const relationships = (...entries: string[]): string => `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.join('')}</Relationships>`

function plainText(node: JSONContent): string {
  return (node.text ?? '') + (node.content ?? []).map(plainText).join(' ')
}

describe('docx from Word', () => {
  it('keeps text in content controls and phonetic guides, lists defined by a list style, and separate lists apart', () => {
    const numbering = `<?xml version="1.0"?><w:numbering NS>
      <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>
      <w:abstractNum w:abstractNumId="2"><w:numStyleLink w:val="Outline"/></w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
      <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
      <w:num w:numId="3"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num></w:numbering>`
    const styles = '<?xml version="1.0"?><w:styles NS><w:style w:type="numbering" w:styleId="Outline"><w:name w:val="Outline"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style></w:styles>'
    const table = `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>
      <w:tr>${tc('Name')}<w:sdt><w:sdtContent>${tc('Filled-in value')}</w:sdtContent></w:sdt></w:tr>
      <w:sdt><w:sdtContent><w:tr>${tc('Repeating')}${tc('row')}</w:tr></w:sdtContent></w:sdt></w:tbl>`
    const ruby = `<w:r><w:ruby><w:rt>${run('かん')}</w:rt><w:rubyBase>${run('漢')}</w:rubyBase></w:ruby></w:r>`
    const body = table + para(run('Before ') + ruby) + listItem('Scope', '2') + listItem('Apples', '1') + listItem('Step one', '3') + listItem('Step two', '3')
    const { content } = readDocx(packageOf(body, { 'word/numbering.xml': numbering, 'word/styles.xml': styles }))

    expect(plainText(content)).toContain('Filled-in value')
    expect(plainText(content)).toContain('Repeating row')
    expect(plainText(content)).toContain('漢')
    expect(plainText(content)).not.toContain('かん')
    expect(content.content?.slice(2).map(node => [node.type, node.content?.length])).toEqual([['orderedList', 1], ['orderedList', 1], ['orderedList', 2]])
  })

  it('reports what it does not keep, so saving over the file asks first', () => {
    const footer = '<?xml version="1.0"?><w:ftr NS><w:p><w:fldSimple w:instr=" DATE "><w:r><w:t>25 September 2026</w:t></w:r></w:fldSimple></w:p></w:ftr>'
    const header = `<?xml version="1.0"?><w:hdr NS>${para(run('Confidential draft'))}</w:hdr>`
    const body = para(run('Cover'), '<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:cols w:num="2"/></w:sectPr>')
      + para('<w:bookmarkStart w:id="0" w:name="terms"/>' + `<w:hyperlink w:anchor="terms">${run('see Terms')}</w:hyperlink>`)
      + para('<w:r><w:rPr><w:b/><w:rPrChange w:id="1" w:author="Ana"><w:rPr/></w:rPrChange></w:rPr><w:t>Made bold</w:t></w:r>')
      + para(`<w:customTag>${run('?')}</w:customTag>`, '<w:bidi/>')
      + '<w:sectPr><w:footerReference w:type="default" r:id="rId2"/></w:sectPr>'
    const rels = relationships(
      '<Relationship Id="rId1" Type="header" Target="header1.xml"/>',
      '<Relationship Id="rId2" Type="footer" Target="footer1.xml"/>',
    )
    const { findings } = readDocx(packageOf(body, { 'word/_rels/document.xml.rels': rels, 'word/header1.xml': header, 'word/footer1.xml': footer }))

    expect(findings.map(finding => finding.construct)).toEqual(expect.arrayContaining([
      'Header and footer text', 'Multiple columns', 'Section breaks', 'Links within the document', 'Bookmarks',
      'Tracked changes', 'Right-to-left paragraphs', 'Other Word content',
    ]))
  })

  it('unpacks only the parts it reads', () => {
    const bytes = zipSync({ 'word/document.xml': strToU8('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:body></w:document>'), 'word/media/unused.bin': new Uint8Array(4096) })
    // Make the unused part's compressed data invalid, so inflating it would throw.
    const at = bytes.findIndex((_, index) => strFromU8(bytes.subarray(index, index + 21)) === 'word/media/unused.bin') + 21
    bytes[at] = 0xff
    expect(() => unzipSync(bytes)).toThrow()
    expect(plainText(readDocx(bytes).content)).toContain('Hi')
  })
})
