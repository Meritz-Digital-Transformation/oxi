# -*- coding: utf-8 -*-
"""Does Word hang a CLOSING BRACKET past the line end, or only 、。，．? (legal__0adfa250 p4)

legal__0adfa250 p4 cell (3686tw, linesAndChars 350/1382, MS Mincho 10.5, compat 15):
「□内職　□その他（　　　　　　）」 -- Word wraps the whole bracket group to line 2
although only the closing 「）」 overflows; Oxi hangs 「）」 and keeps one line.

Sheet: same docDefaults/grid as the source. For each arm a paragraph of N 「あ」 and a
final char X, N swept so that X is the first char that does not fit the line; X in
{、, 。, ）, 」, あ}. Placement body / table cell (3686tw). Readout: the paragraph's line
count (ComputeStatistics(wdStatisticLines) = 1 means X hung).

    python tools/metrics/_pb_hang_bracket_gen.py
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/hang_bracket'); OUT.mkdir(parents=True, exist_ok=True)
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
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
          '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')
SECT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>'
        '<w:cols w:space="425"/><w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="1382"/></w:sectPr>')
XS = ['、', '。', '）', '」', 'あ']


def para(text, snap=True):
    s = '' if snap else '<w:snapToGrid w:val="0"/>'
    return f'<w:p><w:pPr>{s}<w:jc w:val="left"/></w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def body_doc(place, snap):
    ps = []
    for x in XS:
        for n in range(NS[place][0], NS[place][1]):
            ps.append(para('あ' * n + x, snap))
    inner = ''.join(ps)
    if place == 'cell':
        inner = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3686"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3686" w:type="dxa"/></w:tcPr>'
                 + inner + '</w:tc></w:tr></w:tbl><w:p/>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{inner}{SECT}</w:body></w:document>'


NS = {'body': (36, 42), 'cell': (13, 19)}
docs = []
for place in ('body', 'cell'):
    for snap in (1, 0):
        at = OUT / f'{place}_snap{snap}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETTINGS); z.writestr('word/styles.xml', STYLES); z.writestr('word/document.xml', body_doc(place, snap))
        docs.append((place, snap, at))
print('written', OUT)
if '--gen-only' in sys.argv:
    sys.exit(0)

import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
try:
    for place, snap, at in docs:
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            ps = d.Tables(1).Cell(1, 1).Range.Paragraphs if place == 'cell' else d.Paragraphs
            k = 1
            print(f'== {place} snap{snap}: lines per N (N={NS[place][0]}..{NS[place][1]-1})')
            for x in XS:
                row = []
                for n in range(NS[place][0], NS[place][1]):
                    row.append(ps(k).Range.ComputeStatistics(1)); k += 1
                print(f'   X={x}  {row}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
