# -*- coding: utf-8 -*-
"""Where does a HEADER paragraph sit on a `lines` docGrid when the header
distance is smaller than the top margin?
forms__017117b3d0f9921d (JA devG): pgMar top=357 (17.85pt) header=284 (14.2pt),
docGrid lines 360 (18pt), header = one 11pt paragraph (27 ideographic spaces +
受付番号：) with paragraph borders.  Word PDF p4: header rules at 35.85 and 53.88,
header baseline 30.84 -> the header's first line box is 17.85..35.85 = the grid
line that starts AT the top margin; the body's first table rule is 53.88.  Oxi
puts the header line at 35.85..53.85 (one grid line lower) and the table at
72.35, so every page starts 18pt late and section 6 spills to a 9th page.
Arms (A4, lines 360, Normal 11pt ＭＳ 明朝, header = 'HEADER'):
  top357_hdr284  top margin 17.85 > header 14.2   (the real case)
  top357_hdr357  equal
  top357_hdr100  header 5pt, well above the margin
  top1000_hdr284 top margin 50pt > header 14.2 (header well inside the margin)
  top357_hdr284_nogrid  no docGrid (control)
Readout: Word PDF header baseline + first body baseline; Oxi dump header/body y.
Usage: python tools/metrics/_pb_hdrgrid_gen.py   (writes tests/fixtures/hdrgrid/*.docx)
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/hdrgrid'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + W + '>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + W + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
          '<w:kern w:val="2"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')
HEADER = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ' + W + '>'
          '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>HEADER 受付番号：</w:t></w:r></w:p></w:hdr>')


def para(text):
    return '<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % text


def doc(top, hdr, grid):
    body = ''.join(para('本文 %d 行目。' % (i + 1)) for i in range(3))
    sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rId4"/><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="%d" w:right="1134" w:bottom="567" w:left="1134" w:header="%d" w:footer="284" w:gutter="0"/>'
            '<w:cols w:space="425"/>%s</w:sectPr>' % (top, hdr, '<w:docGrid w:type="lines" w:linePitch="360"/>' if grid else ''))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + W + '><w:body>' + body + sect + '</w:body></w:document>')


def write(path, xml):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
        z.writestr('word/_rels/document.xml.rels', DR); z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/styles.xml', STYLES); z.writestr('word/header1.xml', HEADER); z.writestr('word/document.xml', xml)


ARMS = {'top357_hdr284': (357, 284, True), 'top357_hdr357': (357, 357, True), 'top357_hdr100': (357, 100, True),
        'top1000_hdr284': (1000, 284, True), 'top357_hdr284_nogrid': (357, 284, False)}

if __name__ == '__main__':
    for name, (top, hdr, grid) in ARMS.items():
        write(OUT / f'{name}.docx', doc(top, hdr, grid))
    print('wrote', len(ARMS), 'arms to', OUT)
