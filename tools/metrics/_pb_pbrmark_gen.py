# -*- coding: utf-8 -*-
"""Where does the block AFTER a page-break-only paragraph land on the new page?

reports__0079718f (compat 15): paragraph 1612 is `<w:p><w:pPr><w:spacing w:after="200"
w:line="276"/>...</w:pPr><w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:br w:type="page"/></w:r></w:p>`
and the next block (a table) sits at 133.2 on the new page = top margin 123.3 + 9.9.
Oxi puts it at 143.1 = 123.3 + 19.8. Word truth (COM, collapsed start) puts 1612 on the
break's page. Two readings fit Word's 9.9: (i) the paragraph MARK makes one empty line on
the new page and the after is dropped, (ii) no mark line, the after=200 (10pt) carries.

Sheet: 'A' / page-break paragraph (mark size MSZ, after AFTER) / 'B' (Arial 8, before 0).
Arms: AFTER in {0, 200, 400} x MSZ in {16, 40} x compat {15, 14} x splitPgBreakAndParaMark.
Readout: Information(6) of B on page 2 minus the top margin.

    python tools/metrics/_pb_pbrmark_gen.py            # write docx + Word COM readout
    python tools/metrics/_pb_pbrmark_gen.py --gen-only # write docx only (Word busy)
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/pbrmark'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
RPR = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="16"/></w:rPr>'


def settings(compat, split):
    flag = '<w:splitPgBreakAndParaMark/>' if split else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:compat>{flag}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')


def document(after, msz, next_table):
    a = f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>{RPR}</w:pPr><w:r>{RPR}<w:t>A</w:t></w:r></w:p>'
    brk = (f'<w:p><w:pPr><w:spacing w:after="{after}" w:line="276" w:lineRule="auto"/><w:rPr><w:sz w:val="{msz}"/></w:rPr></w:pPr>'
           f'<w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:br w:type="page"/></w:r></w:p>')
    bp = f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>{RPR}</w:pPr><w:r>{RPR}<w:t>B</w:t></w:r></w:p>'
    if next_table:
        b = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'
             f'<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{bp}</w:tc></w:tr></w:tbl><w:p/>')
    else:
        b = bp
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="2466" w:right="2098" w:bottom="2466" w:left="2098" w:header="1814" w:footer="1814" w:gutter="0"/>'
            '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{a}{brk}{b}{sect}</w:body></w:document>'


ARMS = []
for compat, split in ((15, False), (15, True), (14, False)):
    for after in (0, 200, 400):
        for msz in (16, 40):
            for nt in (False, True):
                ARMS.append((compat, split, after, msz, nt))


def name(compat, split, after, msz, nt):
    return f'c{compat}_split{int(split)}_after{after}_msz{msz}_{"tbl" if nt else "par"}.docx'


for arm in ARMS:
    at = OUT / name(*arm)
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
        z.writestr('word/settings.xml', settings(arm[0], arm[1])); z.writestr('word/document.xml', document(arm[2], arm[3], arm[4]))
if '--gen-only' in sys.argv:
    print(f'{len(ARMS)} docx written to {OUT}'); sys.exit(0)

import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for arm in ARMS:
        at = OUT / name(*arm)
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            n = d.Paragraphs.Count
            # B is paragraph 3 (par arm) or the table cell paragraph (tbl arm)
            rng = d.Tables(1).Cell(1, 1).Range if arm[4] else d.Paragraphs(3).Range
            pg = d.Range(rng.Start, rng.Start).Information(3)
            y = round(d.Range(rng.Start, rng.Start).Information(6), 2)
            pbrk = d.Range(d.Paragraphs(2).Range.Start, d.Paragraphs(2).Range.Start).Information(3)
            pmark = d.Range(d.Paragraphs(2).Range.End - 1, d.Paragraphs(2).Range.End - 1).Information(3)
            print(f'{name(*arm):40s} B_page={pg} B_y={y} B_y-margin={round(y - 123.3, 2)} brk_para_start_page={pbrk} mark_page={pmark} pages={d.ComputeStatistics(2)}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
