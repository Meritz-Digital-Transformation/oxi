# -*- coding: utf-8 -*-
"""Does a Latin document's tblHeader row stay alone at a page bottom when the
first data row does not fit? (S1579)

reports__0079718f p181: a tblHeader row («Year | Performance measures |
Expected performance results», trHeight 181 atLeast) fits the page bottom, its
3-cell data row does not; Word starts p182 with both. The CJK rule (S1428,
_pb_keepnext_hdr_gen.py) says the header moves with the data row.

Sheet: Letter, margins 1440, Calibri 11 (no grid), N filler paragraphs, then a
table: row 0 tblHeader (one line), rows 1..3 four lines each, 3 cells.
Arms: header on/off x data rows cantSplit on/off, N swept so the table start
walks over the page bottom. Readout: Information(3) page of row 0 and row 1.

    python tools/metrics/_pb_hdrchain_latin_gen.py
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import os
OUT = Path('tests/fixtures/hdrchain_latin') / os.environ.get('VARIANT', ''); OUT.mkdir(parents=True, exist_ok=True)
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
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="' + os.environ.get('COMPAT', '15') + '"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Times New Roman"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/></w:sectPr>')
B = ('<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
     '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
     '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>')


def p(t):
    return f'<w:p><w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>'


BORDER = os.environ.get('BORDER', '1') == '1'   # 0: no borders at all (technical__00549a8f)
CELLS = int(os.environ.get('CELLS', '3'))        # 1: one cell per row spanning the grid
NLINES = int(os.environ.get('NLINES', '4'))      # lines per data-row cell (one paragraph each)


def table(hdr, cant):
    rows = ''
    for i in range(4):
        n = 1 if i == 0 else NLINES
        trpr = ('<w:tblHeader/>' if hdr else '') if i == 0 else ('<w:cantSplit/>' if cant else '')
        if CELLS == 1:
            cells = '<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/><w:gridSpan w:val="3"/></w:tcPr>' + ''.join(p(f'row {i} cell 0 line {k + 1}') for k in range(n)) + '</w:tc>'
        else:
            cells = ''.join('<w:tc><w:tcPr><w:tcW w:w="3120" w:type="dxa"/></w:tcPr>' + ''.join(p(f'row {i} cell {c} line {k + 1}') for k in range(n)) + '</w:tc>' for c in range(3))
        rows += f'<w:tr><w:trPr>{trpr}</w:trPr>{cells}</w:tr>'
    return (f'<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>{B if BORDER else ""}</w:tblPr><w:tblGrid><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/></w:tblGrid>{rows}</w:tbl>')


def document(n, hdr, cant):
    body = ''.join(p(f'filler line {i + 1}') for i in range(n)) + table(hdr, cant) + p('after the table')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{SECT}</w:body></w:document>'


ARMS = [(h, c, n) for h in tuple(int(v) for v in os.environ.get('HDRS', '1,0').split(',')) for c in (0, 1) for n in range(int(os.environ.get('N0', '44')), int(os.environ.get('N1', '50')))]


def name(h, c, n):
    return f'hdr{h}_cant{c}_n{n:02d}.docx'


for arm in ARMS:
    with zipfile.ZipFile(OUT / name(*arm), 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
        z.writestr('word/settings.xml', SETTINGS); z.writestr('word/styles.xml', STYLES); z.writestr('word/document.xml', document(arm[2], arm[0], arm[1]))
print(f'{len(ARMS)} docx written to {OUT}')
if '--gen-only' in sys.argv:
    sys.exit(0)

import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
try:
    for arm in ARMS:
        d = app.Documents.Open(str((OUT / name(*arm)).resolve()), ReadOnly=True)
        try:
            t = d.Tables(1)
            def at(r, k=1):
                c = t.Cell(r, k).Range; cc = d.Range(c.Start, c.Start)
                return int(cc.Information(3)), cc.Information(6)
            p0, y0 = at(1); p1, y1 = at(2)
            c2 = t.Cell(2, 1).Range; last = d.Range(c2.End - 1, c2.End - 1)
            print(f'{name(*arm)[:-5]:18s} row0 p{p0} y={y0:6.2f}  row1 p{p1} y={y1:6.2f}  row1 last line p{int(last.Information(3))}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
