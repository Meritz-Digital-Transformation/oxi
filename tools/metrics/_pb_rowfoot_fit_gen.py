# -*- coding: utf-8 -*-
"""Does a table row's page-bottom fit include its BOTTOM border width when the row is
not the table's last row?

technical__014819 p4: 12 single-line rows (trHeight 300 atLeast, tcBorders sz=4 = 0.5pt,
row pitch 15.5). Row 11's top is 704.4 in Word's PDF with the body bottom at 720.0
(15.6pt of room for a 15.5pt row) and Word moves row 11 to page 5; Oxi keeps it
(704.25 + 15.5 = 719.75 <= 720). S1482 only adds the foot for the table's LAST row.

Sheet: one exact-height paragraph (lineRule exact = LINE) then a 14-row table of
single-line Calibri 9pt cells, trHeight 300 atLeast, cell borders sz 4. Row 11's top
= 72 + LINE + 11 x 15.5; LINE is chosen so the room left for row 11 is SLACK pt
{15.0, 15.4, 15.6, 15.8, 16.0, 16.2, 16.6}. Readout: page of row 11's first cell.

    python tools/metrics/_pb_rowfoot_fit_gen.py
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/rowfoot_fit'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
RPR = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="18"/></w:rPr>'
BRD_T = '<w:tcBorders><w:top w:val="single" w:sz="{SZ}" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="{SZ}" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="{SZ}" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="{SZ}" w:space="0" w:color="auto"/></w:tcBorders>'


def document(line_pt, nrows=14, cant_split=True, sz=4):
    BRD = BRD_T.replace('{SZ}', str(sz))
    lead = f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{int(round(line_pt * 20))}" w:lineRule="exact"/>{RPR}</w:pPr><w:r>{RPR}<w:t>lead</w:t></w:r></w:p>'
    rows = ''
    for r in range(nrows):
        cells = ''.join(f'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>{BRD}</w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>{RPR}</w:pPr><w:r>{RPR}<w:t>R{r}C{c}</w:t></w:r></w:p></w:tc>' for c in range(4))
        rows += f'<w:tr><w:trPr>{"<w:cantSplit/>" if cant_split else ""}<w:trHeight w:val="300"/></w:trPr>{cells}</w:tr>'
    tbl = f'<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid>{"<w:gridCol w:w=\"2000\"/>" * 4}</w:tblGrid>{rows}</w:tbl>'
    sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{lead}{tbl}<w:p/>{sect}</w:body></w:document>'


SLACKS = (15.5, 16.0, 16.5, 17.0, 17.5, 18.0, 18.5, 19.0)
ARMS = [(s, sz) for sz in (12, 24) for s in SLACKS]
import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    for slack, sz in ARMS:
        cs = True; bw = sz / 8.0
        # row 11 top = 72 + LINE + 11*15.5 ; room = 720 - top = slack  =>  LINE = 720 - slack - 72 - 170.5
        line = 720.0 - slack - 72.0 - 11 * (15.0 + bw)
        at = OUT / f'sz{sz}_slack{slack:.1f}.docx'
        with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
            z.writestr('word/settings.xml', SETT); z.writestr('word/document.xml', document(line, cant_split=cs, sz=sz))
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            t = d.Tables(1)
            ys = []
            for r in (1, 11, 12):
                rg = t.Cell(r, 1).Range; c = d.Range(rg.Start, rg.Start)
                ys.append((r, c.Information(3), round(c.Information(6), 2)))
            print(f'sz={sz} bw={bw:.2f} slack={slack:4.1f} lead={line:.2f} rows(1,11,12)={ys} pages={d.ComputeStatistics(2)}', flush=True)
        finally:
            d.Close(False)
finally:
    app.Quit()
