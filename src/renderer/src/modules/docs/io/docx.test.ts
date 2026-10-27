// @vitest-environment happy-dom
import { getSchema, type JSONContent } from '@tiptap/core'
import { CommentRangeEnd, CommentRangeStart, CommentReference, Document, Header, Packer, Paragraph, TextRun } from 'docx'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import type { HeadersAndFooters } from '../header-footer'
import { documentExtensions } from '../schema'
import { defaultTheme, type DocumentTheme, type PageSetup } from '../theme'
import { sameDocument } from './compare'
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

const page: PageSetup = { width: 16_838, height: 11_906, marginTop: 1_080, marginRight: 1_440, marginBottom: 1_080, marginLeft: 1_800 }
const theme: DocumentTheme = { ...defaultTheme, normal: { ...defaultTheme.normal, fontFamily: 'Georgia', fontSize: 12, lineHeight: 1.5 } }
const small = { type: 'textStyle', attrs: { fontSize: '9pt' } }
const headerFooter: HeadersAndFooters = {
  header: [{ type: 'paragraph', attrs: { textAlign: 'right' }, content: [text('Quarterly report', 'bold'), text(' · draft', 'italic')] }],
  footer: [
    { type: 'paragraph', content: [text('Confidential', 'underline')] },
    { type: 'paragraph', attrs: { textAlign: 'center' }, content: [text('Page ', small), { type: 'pageField', attrs: { kind: 'page' }, marks: [small] }, text(' of ', small), { type: 'pageField', attrs: { kind: 'total' }, marks: [small] }] },
  ],
  firstHeader: [{ type: 'paragraph', content: [text('Ana Lima', { type: 'textStyle', attrs: { fontFamily: 'Verdana', color: '#c00000' } })] }],
  firstFooter: [],
  differentFirstPage: true,
}

/** Everything the model holds, as Sumi would save it. */
const document: JSONContent = {
  type: 'doc',
  attrs: { page, theme, headerFooter, properties: { title: 'Quarterly report', subject: '', creator: 'Ana', keywords: 'sales', description: '' } },
  content: [
    { type: 'title', attrs: { textAlign: 'center' }, content: [text('Quarterly report')] },
    { type: 'tableOfContents', attrs: { levels: 2 } },
    { type: 'heading', attrs: { level: 1 }, content: [text('Summary')] },
    { type: 'heading', attrs: { level: 2 }, content: [text('Regions')] },
    { type: 'heading', attrs: { level: 3 }, content: [text('North')] },
    { type: 'heading', attrs: { level: 4 }, content: [text('Coast')] },
    { type: 'heading', attrs: { level: 5 }, content: [text('Harbour')] },
    { type: 'heading', attrs: { level: 6 }, content: [text('Pier')] },
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
    expect(sameDocument(schema, read.content, document)).toBe(true)
  })

  it('writes a table of contents as a Word TOC field whose entries link to the headings', async () => {
    const xml = strFromU8(unzipSync((await writeDocx(document)).bytes)['word/document.xml'] ?? new Uint8Array())
    expect(xml).toContain(String.raw`<w:instrText>TOC \o &quot;1-2&quot; \h \z \u</w:instrText>`)
    expect(xml).toMatch(/<w:hyperlink [^>]*w:anchor="_Toc1"[^>]*>.*?Summary.*?<\/w:hyperlink>/)
    expect(xml).toContain('<w:bookmarkStart w:name="_Toc1" w:id="1"/>')
    expect(xml).toContain('<w:bookmarkStart w:name="_Toc2" w:id="2"/>')
    // Heading 3 and below are not listed at levels 1 to 2.
    expect(xml).not.toContain('w:anchor="_Toc3"')
  })

  it('lists comments as lost, with counts and what happens to them, and keeps the header', async () => {
    const source = new Document({
      comments: { children: [0, 1].map(id => ({ id, author: 'Reviewer', date: new Date(0), children: [new Paragraph('Check this')] })) },
      sections: [{
        headers: { default: new Header({ children: [new Paragraph('Confidential')] }) },
        children: [0, 1].map(id => new Paragraph({
          children: [new CommentRangeStart(id), new TextRun('Reviewed text'), new CommentRangeEnd(id), new TextRun({ children: [new CommentReference(id)] })],
        })),
      }],
    })
    const { content, findings } = readDocx(new Uint8Array(await Packer.toArrayBuffer(source)))

    expect(findings).toEqual([{ construct: 'Comments', severity: 'dropped', location: '2 comments', suggestedAlternative: 'Not shown, and not kept when saved.' }])
    expect(content.attrs?.['headerFooter']).toMatchObject({ header: [{ type: 'paragraph', content: [{ text: 'Confidential' }] }], footer: [] })
  })

  it('numbers each nested list from its own start, not on from the nested list before it', async () => {
    const numbered = (...values: string[]): JSONContent => ({ type: 'orderedList', content: values.map(value => item(paragraph(text(value)))) })
    const nested: JSONContent = { type: 'doc', content: [{ type: 'bulletList', content: [item(paragraph(text('A')), numbered('x', 'y')), item(paragraph(text('B')), numbered('z'))] }] }
    expect(sameDocument(schema, readDocx((await writeDocx(nested)).bytes).content, nested)).toBe(true)
  })

  it('keeps the link on a linked picture', async () => {
    const linked: JSONContent = { type: 'doc', content: [paragraph({ type: 'image', attrs: { src: png, width: 40, height: 20 }, marks: [{ type: 'link', attrs: { href: 'https://example.com/' } }] })] }
    expect(sameDocument(schema, readDocx((await writeDocx(linked)).bytes).content, linked)).toBe(true)
  })

  it('writes the tabs in a code block as Word tabs, which show as tabs in Word', async () => {
    const code: JSONContent = { type: 'doc', content: [{ type: 'codeBlock', content: [text('build:\n\tgo build')] }] }
    const { bytes } = await writeDocx(code)
    const xml = strFromU8(unzipSync(bytes)['word/document.xml'] ?? new Uint8Array())
    expect(xml).toContain('<w:tab/>')
    expect(xml).not.toMatch(/<w:t[^>]*>[^<]*\t/)
    expect(sameDocument(schema, readDocx(bytes).content, code)).toBe(true)
  })

  it('opens a table inside a table cell without an empty line after it, and keeps an empty line the user typed there', async () => {
    const inner: JSONContent = { type: 'table', content: [{ type: 'tableRow', content: [cell('tableCell', 'Inner')] }] }
    const outer = (...content: JSONContent[]): JSONContent => ({ type: 'doc', content: [{ type: 'table', content: [{ type: 'tableRow', content: [{ type: 'tableCell', content }] }] }] })
    const endsWithTable = outer(paragraph(text('Outer')), inner)
    expect(sameDocument(schema, readDocx((await writeDocx(endsWithTable)).bytes).content, endsWithTable)).toBe(true)
    const emptyAfter = outer(paragraph(text('Outer')), inner, { type: 'paragraph' })
    expect(sameDocument(schema, readDocx((await writeDocx(emptyAfter)).bytes).content, emptyAfter)).toBe(true)
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

  it('restarts a sub-list under each item of a list that restarts its numbering', () => {
    const numbering = `<?xml version="1.0"?><w:numbering NS>
      <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/></w:lvl></w:abstractNum>
      <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
      <w:num w:numId="2"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num></w:numbering>`
    const at = (value: string, numId: string, level: number): string => para(run(value), `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>`)
    const body = at('One', '1', 0) + at('Sub a', '1', 1) + at('Sub b', '1', 1) + para(run('Between')) + at('Restarted', '2', 0) + at('Sub again', '2', 1)
    const { content } = readDocx(packageOf(body, { 'word/numbering.xml': numbering }))

    const restarted = content.content?.[2]
    expect(restarted?.attrs).toBeUndefined()
    expect(restarted?.content?.[0]?.content?.[1]).toMatchObject({ type: 'orderedList', content: [{ content: [{ content: [{ text: 'Sub again' }] }] }] })
    expect(restarted?.content?.[0]?.content?.[1]?.attrs).toBeUndefined()
  })

  it('keeps the spacing and indent of No Spacing and List Paragraph paragraphs', () => {
    const style = (id: string, name: string, pPr: string): string => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:pPr>${pPr}</w:pPr></w:style>`
    const styles = `<?xml version="1.0"?><w:styles NS><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:style>`
      + style('NoSpacing', 'No Spacing', '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>') + style('ListParagraph', 'List Paragraph', '<w:ind w:left="720"/>') + '</w:styles>'
    const { content, findings } = readDocx(packageOf(para(run('Address line'), '<w:pStyle w:val="NoSpacing"/>') + para(run('Indented'), '<w:pStyle w:val="ListParagraph"/>'), { 'word/styles.xml': styles }))

    expect(findings).toEqual([])
    expect(content.content?.map(node => node.attrs)).toEqual([{ spaceAfter: 0, lineHeight: 1 }, { indentLeft: 36 }])
  })

  it('reports links to files and other apps, which it does not keep', () => {
    const rels = relationships('<Relationship Id="rId1" Type="hyperlink" Target="file:///Users/ana/plan.xlsx" TargetMode="External"/>')
    const field = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "onenote:///Notes/Plan.one" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
      + run('notes') + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    for (const body of [para(`<w:hyperlink r:id="rId1">${run('the plan')}</w:hyperlink>`), para(field), para(`<w:fldSimple w:instr=" HYPERLINK &quot;file:///x.docx&quot; ">${run('x')}</w:fldSimple>`)]) {
      const { findings } = readDocx(packageOf(body, { 'word/_rels/document.xml.rels': rels }))
      expect(findings.map(finding => finding.construct)).toEqual(['Links to files and other apps'])
    }
  })

  it('reports cropped, rotated and flipped pictures, which show whole and upright', () => {
    const rels = relationships('<Relationship Id="rId1" Type="image" Target="media/image1.png"/>')
    const picture = (srcRect: string, xfrm: string): string => '<w:r><w:drawing><wp:inline xmlns:wp="wp" xmlns:a="a" xmlns:pic="pic"><wp:extent cx="381000" cy="190500"/>'
      + `<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId1"/>${srcRect}</pic:blipFill><pic:spPr><a:xfrm${xfrm}/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
    const media = Uint8Array.from(atob(png.split(',')[1] ?? ''), character => character.charCodeAt(0))
    const open = (body: string): readonly string[] => {
      const bytes = unzipSync(packageOf(body, { 'word/_rels/document.xml.rels': rels }))
      return readDocx(zipSync({ ...bytes, 'word/media/image1.png': media })).findings.map(finding => finding.construct)
    }

    expect(open(para(picture('<a:srcRect/>', ' rot="0"')))).toEqual([])
    for (const changed of [picture('<a:srcRect l="10000" r="5000"/>', ''), picture('', ' rot="5400000"'), picture('', ' flipH="1"')]) {
      expect(open(para(changed))).toEqual(['Picture cropping and rotation'])
    }
  })

  it('keeps the link Word stores on a picture itself, and reports one it cannot keep', () => {
    const rels = relationships(
      '<Relationship Id="rId1" Type="image" Target="media/image1.png"/>',
      '<Relationship Id="rId2" Type="hyperlink" Target="https://example.com/" TargetMode="External"/>',
      '<Relationship Id="rId3" Type="hyperlink" Target="file:///Users/ana/plan.xlsx" TargetMode="External"/>',
    )
    const picture = (id: string): string => '<w:r><w:drawing><wp:inline xmlns:wp="wp" xmlns:a="a" xmlns:pic="pic"><wp:extent cx="381000" cy="190500"/>'
      + `<wp:docPr id="1" name="Picture 1"><a:hlinkClick r:id="${id}"/></wp:docPr>`
      + '<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    const media = Uint8Array.from(atob(png.split(',')[1] ?? ''), character => character.charCodeAt(0))
    const open = (id: string): ReturnType<typeof readDocx> => {
      const bytes = unzipSync(packageOf(para(picture(id)), { 'word/_rels/document.xml.rels': rels }))
      return readDocx(zipSync({ ...bytes, 'word/media/image1.png': media }))
    }

    const linked = open('rId2')
    expect(linked.findings).toEqual([])
    expect(linked.content.content?.[0]?.content?.[0]).toMatchObject({ type: 'image', marks: [{ type: 'link', attrs: { href: 'https://example.com/' } }] })
    expect(open('rId3').findings.map(finding => finding.construct)).toEqual(['Links to files and other apps'])
  })

  it('reports right indents, which it does not keep', () => {
    const open = (pPr: string): readonly string[] => readDocx(packageOf(para(run('Indented'), pPr))).findings.map(finding => finding.construct)
    expect(open('<w:ind w:left="720" w:right="0"/>')).toEqual([])
    expect(open('<w:ind w:left="720" w:right="720"/>')).toEqual(['Right indents'])
    expect(open('<w:ind w:end="360"/>')).toEqual(['Right indents'])
  })

  it('reports table alignment, row heights and cell alignment, which it does not keep', () => {
    const open = (tblPr: string, trPr: string, tcPr: string): readonly string[] => {
      const body = `<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:trPr>${trPr}</w:trPr><w:tc><w:tcPr>${tcPr}</w:tcPr>${para(run('Cell'))}</w:tc></w:tr></w:tbl>`
      return readDocx(packageOf(body)).findings.map(finding => finding.construct)
    }
    expect(open('<w:jc w:val="left"/>', '', '<w:vAlign w:val="top"/>')).toEqual([])
    for (const changed of [open('<w:jc w:val="center"/>', '', ''), open('', '<w:trHeight w:val="900"/>', ''), open('', '', '<w:vAlign w:val="center"/>')]) {
      expect(changed).toEqual(['Table layout'])
    }
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
      'Fields', 'Multiple columns', 'Section breaks', 'Links within the document', 'Bookmarks',
      'Tracked changes', 'Right-to-left paragraphs', 'Other Word content',
    ]))
  })

  it('reads page number fields and a different first page from a Word header and footer', () => {
    const styles = '<?xml version="1.0"?><w:styles NS><w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:pPr><w:spacing w:after="0"/></w:pPr></w:style></w:styles>'
    const field = (code: string, shown: string): string => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r>`
      + `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run(shown)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`
    const header = `<?xml version="1.0"?><w:hdr NS>${para(run('Page ') + field('PAGE \\* MERGEFORMAT', '3') + run(' of ') + '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple>', '<w:pStyle w:val="Header"/><w:jc w:val="right"/>')}</w:hdr>`
    const cover = `<?xml version="1.0"?><w:ftr NS>${para(run('Cover'))}</w:ftr>`
    const body = para(run('Text')) + '<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="first" r:id="rId2"/><w:titlePg/></w:sectPr>'
    const rels = relationships('<Relationship Id="rId1" Type="header" Target="header1.xml"/>', '<Relationship Id="rId2" Type="footer" Target="footer1.xml"/>')
    const { content, findings } = readDocx(packageOf(body, { 'word/_rels/document.xml.rels': rels, 'word/header1.xml': header, 'word/footer1.xml': cover, 'word/styles.xml': styles }))

    expect(findings).toEqual([])
    expect(content.attrs?.['headerFooter']).toEqual({
      header: [{ type: 'paragraph', attrs: { textAlign: 'right' }, content: [{ type: 'text', text: 'Page ' }, { type: 'pageField', attrs: { kind: 'page' } }, { type: 'text', text: ' of ' }, { type: 'pageField', attrs: { kind: 'total' } }] }],
      footer: [],
      firstHeader: [],
      firstFooter: [{ type: 'paragraph', content: [{ type: 'text', text: 'Cover' }] }],
      differentFirstPage: true,
    })
  })

  it('reads the table of contents Word makes, without its entries', () => {
    const field = (code: string): string => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`
    const entry = (title: string, bookmark: string): string => `<w:hyperlink w:anchor="${bookmark}" w:history="1">${run(title)}<w:r><w:tab/></w:r>${field(`PAGEREF ${bookmark} \\h`)}${run('2')}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>`
    const contents = '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>'
      + para(field('TOC \\o "1-4" \\h \\z \\u') + entry('Scope', '_Toc101'), '<w:pStyle w:val="TOC1"/>')
      + para(entry('Terms', '_Toc102'), '<w:pStyle w:val="TOC2"/>')
      + para('<w:r><w:fldChar w:fldCharType="end"/></w:r>')
      + '</w:sdtContent></w:sdt>'
    const { content, findings } = readDocx(packageOf(contents + para('<w:bookmarkStart w:id="0" w:name="_Toc101"/>' + run('Scope') + '<w:bookmarkEnd w:id="0"/>')))

    expect(findings).toEqual([])
    expect(content.content).toEqual([{ type: 'tableOfContents', attrs: { levels: 4 } }, { type: 'paragraph', content: [{ type: 'text', text: 'Scope' }] }])
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
