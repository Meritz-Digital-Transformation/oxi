# -*- coding: utf-8 -*-
"""Which face sets the height of an EMPTY paragraph whose line spacing is a
MULTIPLE (line=276/360/480 auto) in a Latin document -- the mark's ASCII face
(S989/S1382: measured for single spacing only) or the eastAsia face Oxi's
S989 gate falls back to when the multiplier is not 1.0?

legal__0030f893 (en-ZA, theme ea "" -> Jpan ＭＳ 明朝, compat 15): in a table
cell, empty paragraphs with `line=360 auto` and an Arial 12pt / 11pt mark
measure 21.0 / 18.75 in Word (COM Information(6) steps = Arial hhea 13.8 /
12.65 x 1.5) while Oxi advances 23.35 / 21.4 (MS Mincho 83/64 x 1.5): +2.4pt
per empty line, six of them on page 3, and the page's last paragraph moves.

Sheet: a body paragraph "A", the probe (empty), "B"; then a 1x1 table whose
cell holds "A", the probe, "B". Arms: line {240, 276, 360, 480} x mark
{Arial 10, Arial 12, Calibri 11} x mark eastAsia {none, ＭＳ 明朝}.
Readout: Information(6) of the probe and of the paragraph after it, body and
cell (collapsed starts), so height = y(next) - y(probe).

    python tools/metrics/_pb_emptymult_face_gen.py
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/emptymult_face'); OUT.mkdir(parents=True, exist_ok=True)
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
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="ＭＳ 明朝" w:hAnsi="Calibri" w:cs="Times New Roman"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-ZA" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>')


def rpr(face, sz, ea):
    fonts = f'<w:rFonts w:ascii="{face}" w:hAnsi="{face}" w:cs="{face}"' + (f' w:eastAsia="{ea}"' if ea else '') + '/>'
    return f'<w:rPr>{fonts}<w:sz w:val="{sz}"/></w:rPr>'


def para(text, line, face, sz, ea):
    r = rpr(face, sz, ea)
    run = f'<w:r>{r}<w:t xml:space="preserve">{text}</w:t></w:r>' if text else ''
    return f'<w:p><w:pPr><w:spacing w:line="{line}" w:lineRule="auto"/>{r}</w:pPr>{run}</w:p>'


def document(line, face, sz, ea):
    trio = para('A', line, face, sz, ea) + para('', line, face, sz, ea) + para('B', line, face, sz, ea)
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
           f'<w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr>{trio}</w:tc></w:tr></w:tbl>')
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="425"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{trio}{tbl}<w:p/>{sect}</w:body></w:document>'


MARKS = [('Arial', 20, 'a10'), ('Arial', 24, 'a12'), ('Calibri', 22, 'c11')]
ARMS = [(line, face, sz, tag, ea, etag) for line in (240, 276, 360, 480) for face, sz, tag in MARKS for ea, etag in (('', 'none'), ('ＭＳ 明朝', 'min'))]


def name(line, face, sz, tag, ea, etag):
    return f'l{line}_{tag}_{etag}.docx'


for arm in ARMS:
    at = OUT / name(*arm)
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
        z.writestr('word/settings.xml', SETTINGS); z.writestr('word/styles.xml', STYLES); z.writestr('word/document.xml', document(arm[0], arm[1], arm[2], arm[4]))
print(f'{len(ARMS)} docx written to {OUT}')
if '--gen-only' in sys.argv:
    sys.exit(0)

import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
try:
    print(f'{"arm":18s} {"body":>7s} {"cell":>7s}   (hhea x mult: Arial10 11.5 / Arial12 13.8 / Calibri11 13.4; Mincho 83/64: 12.97 / 15.56 / 14.27)')
    for arm in ARMS:
        at = OUT / name(*arm)
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            def y(i):
                r = d.Paragraphs(i).Range; return d.Range(r.Start, r.Start).Information(6)
            body = y(3) - y(2)
            c = d.Tables(1).Cell(1, 1).Range
            ps = [p for p in c.Paragraphs]
            cy = [d.Range(p.Range.Start, p.Range.Start).Information(6) for p in ps[:3]]
            cell = cy[2] - cy[1]
            print(f'{name(*arm)[:-5]:18s} {body:7.2f} {cell:7.2f}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
