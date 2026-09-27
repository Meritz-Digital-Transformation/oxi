# -*- coding: utf-8 -*-
"""A table with THREE tblHeader rows starting at a page bottom where only the first
one or two header rows fit (S1587; reference__0096d8c9 p155/156).

Sheet: Times New Roman 11, Letter, N filler lines, then a 3-column table: rows 0-2
tblHeader (row 2 three lines tall), then 12 data rows. N swept so the table start
walks over the page bottom. Readout: page and y of each header row's first cell and of
data row 3, from the Word PDF: the header labels H0/H1/H2 are unique strings, and on a
continuation page Word repeats them.

    python tools/metrics/_pb_hdr3_split_gen.py
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/hdr3_split'); OUT.mkdir(parents=True, exist_ok=True)
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
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="' + os.environ.get('COMPAT', '14') + '"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
B = ('<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
     '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>')


def p(t):
    return f'<w:p><w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>'


def row(label, n, hdr):
    cells = ''.join('<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>' + ''.join(p(f'{label} c{c} l{k + 1}') for k in range(n)) + '</w:tc>' for c in range(3))
    return f'<w:tr><w:trPr><w:cantSplit/>{"<w:tblHeader/>" if hdr else ""}</w:trPr>{cells}</w:tr>'


def document(n):
    rows = row('HZERO', 1, True) + row('HONE', 1, True) + row('HTWO', 3, True) + ''.join(row(f'DATA{i}', 1, False) for i in range(int(os.environ.get('NDATA', '12'))))
    tbl = f'<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>{B}</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>{rows}</w:tbl>'
    body = ''.join(p(f'filler {i}') for i in range(n)) + tbl + p('after')
    sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


NS = list(range(int(os.environ.get('N0', '41')), int(os.environ.get('N1', '46'))))
for n in NS:
    with zipfile.ZipFile(OUT / f'n{n}.docx', 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
        z.writestr('word/settings.xml', SETTINGS); z.writestr('word/styles.xml', STYLES); z.writestr('word/document.xml', document(n))
print('written', OUT)
if '--gen-only' in sys.argv or '--read-only' in sys.argv and False:
    sys.exit(0)
import win32com.client, fitz
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
try:
    for n in NS:
        at = OUT / f'n{n}.docx'; pdf = str((OUT / f'n{n}.pdf').resolve())
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            d.ExportAsFixedFormat(pdf, 17)
        finally:
            d.Close(False)
        doc = fitz.open(pdf); out = []
        for pn in range(len(doc)):
            for b in doc[pn].get_text('dict')['blocks']:
                for l in b.get('lines', []):
                    t = ''.join(s['text'] for s in l['spans'])
                    if 'c0 l1' in t and (t.startswith('H') or t.startswith('DATA0') or t.startswith('DATA1 ') or (pn >= 2 and t.startswith('DATA') and not out[-1].startswith(f'p{pn + 1}:DATA'))):
                        out.append(f'p{pn + 1}:{t.split()[0]}@{l["bbox"][1]:.1f}')
        print(f'n={n}: ' + '  '.join(out), flush=True)
finally:
    app.Quit()
