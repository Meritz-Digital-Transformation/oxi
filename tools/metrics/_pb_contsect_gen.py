# -*- coding: utf-8 -*-
"""Does a CONTINUOUS section break ever start a new page when only the docGrid
charSpace / pgNumType change?  reference__0ea3ec86480140c2 with S1536 on: the
last line of section 1 lands alone on p4, the 3pt paragraph carrying
<w:type w:val="continuous"/> follows it, and Oxi emits an EMPTY p5 before the
next section's content on p6.  Word (PDF) starts that section right after the
previous content (it opened p4 only because p3 was full).  pgSz/pgMar are the
same in both sections; the differences are docGrid charSpace 3042 -> 3194 and
<w:pgNumType w:start="88"/>.
Arms (section 1 = 6 lines, then a 3pt exact paragraph holding the sectPr):
  none  - section 2 identical geometry (control)
  grid  - section 2 charSpace 3194 (section 1 3042)
  pgnum - section 2 pgNumType start=88
  both  - grid + pgnum
  top   - `both`, and section 1 is exactly one page long so the break sits at a
          page top (the real case: the sectPr paragraph is the first thing on p4)
Readout: page count (Word PDF vs Oxi) and the y of section 2's first line.
Usage: python tools/metrics/_pb_contsect_gen.py   (writes tests/fixtures/contsect/*.docx)
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/contsect'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + W + '>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + W + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
          '<w:kern w:val="2"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')
PG = '<w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="1304" w:right="1021" w:bottom="1134" w:left="1021" w:header="680" w:footer="567" w:gutter="0"/>'


def para(text):
    return '<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % text


def sect1(charspace, n_lines):
    body = ''.join(para('第一節の本文 %d 行目。' % (i + 1)) for i in range(n_lines))
    # the 3pt exact paragraph that carries the continuous break (as the real doc)
    body += ('<w:p><w:pPr><w:spacing w:line="60" w:lineRule="exact"/><w:sectPr><w:type w:val="continuous"/>' + PG +
             '<w:cols w:space="440"/><w:docGrid w:type="linesAndChars" w:linePitch="411" w:charSpace="%d"/></w:sectPr></w:pPr></w:p>' % charspace)
    return body


def sect2(charspace, pgnum):
    body = ''.join(para('第二節の本文 %d 行目。' % (i + 1)) for i in range(4))
    body += ('<w:sectPr><w:type w:val="continuous"/>' + PG + ('<w:pgNumType w:start="88"/>' if pgnum else '') +
             '<w:cols w:space="440"/><w:docGrid w:type="linesAndChars" w:linePitch="411" w:charSpace="%d"/></w:sectPr>' % charspace)
    return body


def doc(arm):
    cs2 = 3194 if arm in ('grid', 'both', 'top') else 3042
    pgnum = arm in ('pgnum', 'both', 'top')
    # page: content height = 841.9 - 65.2 - 56.7 = 720; line pitch 20.55 -> 35 lines fill a page.
    # `top`: 35 body lines so the sectPr paragraph is pushed to the top of p2
    n = 35 if arm == 'top' else 6
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + W + '><w:body>' +
            sect1(3042, n) + sect2(cs2, pgnum) + '</w:body></w:document>')


def write(path, xml):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
        z.writestr('word/_rels/document.xml.rels', DR); z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/styles.xml', STYLES); z.writestr('word/document.xml', xml)


if __name__ == '__main__':
    for arm in ('none', 'grid', 'pgnum', 'both', 'top'):
        write(OUT / f'{arm}.docx', doc(arm))
    print('wrote 5 arms to', OUT)
