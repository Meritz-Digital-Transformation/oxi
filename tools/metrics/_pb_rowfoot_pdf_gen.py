# -*- coding: utf-8 -*-
"""Row page-bottom fit vs cell border width, read from Word's PDF (no Info(6) quantization).

Same sheet as _pb_rowfoot_fit_gen.py: exact-height lead paragraph, then a 14-row table of
single-line Calibri 9pt cells (trHeight 300 atLeast, cell borders sz=SZ on all four sides).
The lead height puts row 12's nominal top at 720 - ROOM. Readout from ExportAsFixedFormat:
the horizontal rules on page 1 (count >= 14 => row 12 stayed) and the last rule y.

    python tools/metrics/_pb_rowfoot_pdf_gen.py [--sz 4,12] [--rooms 15.0:17.5:0.1]
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import fitz
OUT = Path('tests/fixtures/rowfoot_pdf'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
RPR = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="18"/></w:rPr>'


def document(line_pt, sz, nrows=14):
    brd = ''.join(f'<w:{side} w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>' for side in ('top', 'left', 'bottom', 'right'))
    lead = f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{int(round(line_pt * 20))}" w:lineRule="exact"/>{RPR}</w:pPr><w:r>{RPR}<w:t>lead</w:t></w:r></w:p>'
    rows = ''
    for r in range(nrows):
        cells = ''.join(f'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:tcBorders>{brd}</w:tcBorders></w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>{RPR}</w:pPr><w:r>{RPR}<w:t>R{r}C{c}</w:t></w:r></w:p></w:tc>' for c in range(4))
        rows += f'<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="300"/></w:trPr>{cells}</w:tr>'
    tbl = f'<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid>{"<w:gridCol w:w=\"2000\"/>" * 4}</w:tblGrid>{rows}</w:tbl>'
    sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{lead}{tbl}<w:p/>{sect}</w:body></w:document>'


def hrules(pdf_path, page=0):
    p = fitz.open(pdf_path)[page]
    ys = sorted({round(d['rect'].y0, 2) for d in p.get_drawings() if d['rect'].width > 100 and d['rect'].height < 4.0})
    return ys


args = sys.argv[1:]
szs = [4, 8, 12, 16, 24]; rooms = None
if '--sz' in args: szs = [int(v) for v in args[args.index('--sz') + 1].split(',')]
if '--rooms' in args:
    a, b, st = [float(v) for v in args[args.index('--rooms') + 1].split(':')]
    rooms = []; v = a
    while v <= b + 1e-9: rooms.append(round(v, 2)); v += st
import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for sz in szs:
        bw = sz / 8.0
        pitch_guess = 15.0 + bw
        rs = rooms or [round(pitch_guess - 0.5 + 0.1 * k, 2) for k in range(0, 26)]
        for room in rs:
            line = 720.0 - room - 72.0 - 11 * pitch_guess
            at = OUT / f'sz{sz}_room{room:.2f}.docx'; pdf = at.with_suffix('.pdf')
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
                z.writestr('word/settings.xml', SETT); z.writestr('word/document.xml', document(line, sz))
            d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
            try:
                d.ExportAsFixedFormat(str(pdf.resolve()), 17)
            finally:
                d.Close(False)
            ys = hrules(str(pdf))
            top11 = ys[11] if len(ys) > 11 else None; top12 = ys[12] if len(ys) > 12 else None
            stays = len(ys) >= 14
            pitch = (ys[11] - ys[1]) / 10 if len(ys) > 11 else None
            print(f'sz={sz:2d} bw={bw:.2f} room={room:5.2f} rules={len(ys):2d} pitch={pitch} row12_top={top12} room_real={None if top12 is None else round(720 - top12, 2)} {"STAY" if stays else "MOVE"}', flush=True)
finally:
    app.Quit()
