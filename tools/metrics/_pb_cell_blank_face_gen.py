# -*- coding: utf-8 -*-
"""Which face sets the height of a TEXTLESS paragraph inside a table cell of a
CJK document -- the eastAsia face (MS Mincho 83/64) or the ascii face (Century)?

legal__0adfa250 p3: a cell of three 8pt paragraphs (snapToGrid=0) -- a
paragraph that holds only a wrapNone text box, one with 「年 月 日」 and one with
a single ASCII space. Word's row is ~29.9 (rule to rule), Oxi's 33.25 =
3 x 10.375 (MS Mincho 8pt 83/64); Century's 8pt line would give ~29.1+pad.

Sheet (docDefaults ascii Century / eastAsia ＭＳ 明朝, Normal sz 22, linesAndChars
350 / charSpace 1382, compat 15): a 1x1 table per arm whose cell holds three
8pt paragraphs: A = content kind, then 「年月日」, then content kind again.
Kinds: empty / one ASCII space / one ideographic space / text 「年」.
snapToGrid 0 / 1. Readout: Information(6) of the paragraph after the table
minus that of the table's first paragraph (row height incl. borders).

    python tools/metrics/_pb_cell_blank_face_gen.py
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/cell_blank_face'); OUT.mkdir(parents=True, exist_ok=True)
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
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
          '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr>'
          '<w:rPr><w:kern w:val="2"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style></w:styles>')
KINDS = {'empty': None, 'asciisp': ' ', 'ideosp': '　', 'text': '年'}


def p(text, snap):
    s = '' if snap else '<w:snapToGrid w:val="0"/>'
    ppr = f'<w:pPr>{s}<w:rPr><w:sz w:val="16"/></w:rPr></w:pPr>'
    run = f'<w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>' if text is not None else ''
    return f'<w:p>{ppr}{run}</w:p>'


def document(kind, snap):
    k = KINDS[kind]
    cell = p(k, snap) + p('　　年　月　日', snap) + p(k, snap)
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
           f'<w:tblGrid><w:gridCol w:w="2551"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2551" w:type="dxa"/></w:tcPr>{cell}</w:tc></w:tr></w:tbl>')
    body = '<w:p><w:r><w:t>前</w:t></w:r></w:p>' + tbl + '<w:p><w:r><w:t>後</w:t></w:r></w:p>'
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:cols w:space="425"/><w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="1382"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


ARMS = [(k, s) for k in KINDS for s in (0, 1)]
for k, s in ARMS:
    with zipfile.ZipFile(OUT / f'{k}_snap{s}.docx', 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
        z.writestr('word/settings.xml', SETTINGS); z.writestr('word/styles.xml', STYLES); z.writestr('word/document.xml', document(k, s))
print(f'{len(ARMS)} docx written to {OUT}')
if '--gen-only' in sys.argv:
    sys.exit(0)

import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
try:
    print('arm              row   p1->p2  p2->p3   (MS Mincho 8pt 83/64 = 10.375)')
    for k, s in ARMS:
        d = app.Documents.Open(str((OUT / f'{k}_snap{s}.docx').resolve()), ReadOnly=True)
        try:
            y = lambda r: d.Range(r.Start, r.Start).Information(6)
            c = d.Tables(1).Cell(1, 1).Range; ps = [q.Range for q in c.Paragraphs]
            after = d.Paragraphs(d.Paragraphs.Count).Range
            print(f'{k:8s} snap{s}  {y(after) - y(ps[0]):6.2f}  {y(ps[1]) - y(ps[0]):6.2f}  {y(ps[2]) - y(ps[1]):6.2f}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
