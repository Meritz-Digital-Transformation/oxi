# -*- coding: utf-8 -*-
"""Does a cell-relative floating picture that straddles the page bottom stop Word
from splitting the row (whole row to the next page), or does the row split anyway?

correspondence__101d483d (JA, compat 14): a 1-row 1-cell table with 7 wrapSquare photos
positioned relativeFrom="margin" (layoutInCell=1). Oxi's S1168 sees the third photo
straddle the page bottom and moves the split to the row top (whole row to the next
page); Word splits the row with three paragraphs on the first page. forms__005a5d91
and educational__00161422 (EN, compat 15) have the same construction with 10pt-tall
pictures straddling by 1.6pt and there Word DOES move the row (S1168 was right).

Sheet: FILL filler lines, then a 1-row 1-cell table whose cell holds 14 one-line
paragraphs (Arial 11) and ONE anchored picture (wrapSquare, layoutInCell=1,
positionH margin left, positionV margin posOffset=OFF) of height H. FILL puts the
row top at ~500pt on a 792pt page (content bottom 720). Arms: H in {10, 80, 160}
x OFF in {150, 200, 215} (picture top = row_top + OFF; OFF 200/215 straddle the
bottom with H >= 80, 215 straddles even for H=10) x compat {14, 15}.
Readout: page of the cell's first and last paragraph (COM, collapsed start).

    python tools/metrics/_pb_cellfloat_split_gen.py
"""
import zipfile, sys, os, struct, zlib
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/cellfloat_split'); OUT.mkdir(parents=True, exist_ok=True)
W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
     'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
     'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>')
RPR = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/></w:rPr>'


def png(w=100, h=100):
    raw = b''.join(b'\x00' + bytes([0x80, 0x80, 0xff] * w) for _ in range(h))
    def chunk(t, d):
        return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b''))


def settings(compat):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')


def para(text):
    return f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>{RPR}</w:pPr><w:r>{RPR}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def anchor(h_pt, off_pt):
    emu = lambda pt: int(round(pt * 12700))
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            f'<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH>'
            f'<wp:positionV relativeFrom="margin"><wp:posOffset>{emu(off_pt)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{emu(120)}" cy="{emu(h_pt)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>'
            f'<wp:docPr id="1" name="Picture 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu(120)}" cy="{emu(h_pt)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            f'</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


def document(fill, h_pt, off_pt, ncell=14):
    body = ''.join(para(f'F{i}') for i in range(fill))
    cell = ''
    for i in range(ncell):
        if i == 1:
            cell += (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>{RPR}</w:pPr>'
                     f'{anchor(h_pt, off_pt)}<w:r>{RPR}<w:t xml:space="preserve">C{i} cell line with the picture anchored here</w:t></w:r></w:p>')
        else:
            cell += para(f'C{i} cell line number {i} of the single row')
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
           f'<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{cell}</w:tc></w:tr></w:tbl>')
    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{tbl}<w:p/>{sect}</w:body></w:document>'


ARMS = [(compat, h, off) for compat in (15, 14) for h in (10, 80, 160) for off in (150, 200, 215)]
FILL = 34  # 34 x 12.65 = 430 -> row top ~502; content bottom 720 -> 218pt left in the row


def name(compat, h, off):
    return f'c{compat}_h{h}_off{off}.docx'


for arm in ARMS:
    at = OUT / name(*arm)
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
        z.writestr('word/settings.xml', settings(arm[0])); z.writestr('word/document.xml', document(FILL, arm[1], arm[2]))
        z.writestr('word/media/image1.png', png())
if '--gen-only' in sys.argv:
    print(f'{len(ARMS)} docx written to {OUT}'); sys.exit(0)

import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for arm in ARMS:
        at = OUT / name(*arm)
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            t = d.Tables(1); c = t.Cell(1, 1).Range
            first = d.Range(c.Start, c.Start); last = d.Range(c.End - 2, c.End - 2)
            p_first = first.Information(3); y_first = round(first.Information(6), 2)
            p_last = last.Information(3); y_last = round(last.Information(6), 2)
            # first paragraph of the cell on page 2 => whole row moved
            print(f'{name(*arm):22s} row_first=p{p_first}@{y_first:7.2f} row_last=p{p_last}@{y_last:7.2f} pages={d.ComputeStatistics(2)} '
                  f'{"ROW_MOVED" if p_first == 2 else "SPLIT" if p_last == 2 else "FITS"}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
