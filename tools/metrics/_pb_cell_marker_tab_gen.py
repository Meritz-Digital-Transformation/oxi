# -*- coding: utf-8 -*-
"""Where does a CELL list paragraph's text start when the number is wider than the hanging?
(S1590; reports__00870bdf table 9)

Sheet: Calibri 10 bold numbering (upperLetter 'A.'...), one-column table, each cell
paragraph ind left=159 hanging=H, defaultTabStop 720. Arms: hanging H in {140,180,200,260}
x letter in {A,B,I,W,M} (different marker widths). Readout: Information(5) x of the first
text character relative to the cell's text start, and the marker width from Calibri Bold.
Word rule under test: marker_w > hanging -> text at the next default stop (36pt multiple
from the cell text start); else at the hanging stop (left).

    python tools/metrics/_pb_cell_marker_tab_gen.py
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/cell_marker_tab'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="20"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
LETTERS = ['A', 'B', 'I', 'W', 'M']
HANGS = [140, 180, 200, 260]


def numbering():
    nums = ''
    for i, L in enumerate(LETTERS):
        # one abstract per letter: start value picks the letter (upperLetter)
        nums += (f'<w:abstractNum w:abstractNumId="{i}"><w:lvl w:ilvl="0"><w:start w:val="{ord(L) - 64}"/><w:numFmt w:val="upperLetter"/>'
                 f'<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="159" w:hanging="180"/></w:pPr>'
                 f'<w:rPr><w:b/></w:rPr></w:lvl></w:abstractNum>')
    # one num instance per (hanging, letter) cell, restarting at the letter, so the
    # letters do not continue from paragraph to paragraph
    k = 0
    for h in HANGS:
        for i, L in enumerate(LETTERS):
            k += 1
            nums += (f'<w:num w:numId="{k}"><w:abstractNumId w:val="{i}"/>'
                     f'<w:lvlOverride w:ilvl="0"><w:startOverride w:val="{ord(L) - 64}"/></w:lvlOverride></w:num>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {W}>{nums}</w:numbering>'


def doc():
    rows = ''
    k = 0
    for h in HANGS:
        for i, L in enumerate(LETTERS):
            k += 1
            p = (f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{k}"/></w:numPr><w:ind w:left="159" w:hanging="{h}"/></w:pPr>'
                 f'<w:r><w:rPr><w:b/></w:rPr><w:t>Private Sector Entities</w:t></w:r></w:p>')
            rows += f'<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{p}</w:tc></w:tr>'
    tbl = f'<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>{rows}</w:tbl><w:p/>'
    sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{tbl}{sect}</w:body></w:document>'


at = OUT / 'cell_marker_tab.docx'
with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
    z.writestr('word/settings.xml', SETTINGS); z.writestr('word/styles.xml', STYLES); z.writestr('word/numbering.xml', numbering()); z.writestr('word/document.xml', doc())
print('written', at)
if '--gen-only' in sys.argv:
    sys.exit(0)
import win32com.client, fitz
pdf = str((OUT / 'cell_marker_tab.pdf').resolve())
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
try:
    d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
    try:
        d.ExportAsFixedFormat(pdf, 17)
    finally:
        d.Close(False)
finally:
    app.Quit()
doc = fitz.open(pdf)
chars = []
for b in doc[0].get_text('rawdict')['blocks']:
    for l in b.get('lines', []):
        for s_ in l['spans']:
            for c in s_['chars']:
                chars.append((round(c['origin'][1], 1), c['c'], c['bbox'][0], c['bbox'][2]))
from collections import defaultdict
rows = defaultdict(list)
for y, c, x0, x1 in chars:
    rows[y].append((x0, x1, c))
k = 0
ys = sorted(rows)
for h in HANGS:
    for L in LETTERS:
        cs = sorted(rows[ys[k]]); k += 1
        dot = [c for c in cs if c[2] == '.'][0]
        p = [c for c in cs if c[2] == 'P'][0]
        mstart = cs[0][0]
        print(f'hanging {h/20:5.2f}  marker {cs[0][2]}. w={dot[1] - mstart:5.2f} (start {mstart:6.2f})  text x={p[0]:7.2f}  {"JUMP" if p[0] > 100 else "hang-stop"}', flush=True)
