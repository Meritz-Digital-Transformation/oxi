# -*- coding: utf-8 -*-
"""On a lines docGrid, does a paragraph whose SNAPPED box overhangs the body bottom stay
on the page when its centred glyph box fits?

administrative__108cf8: Meiryo 12pt on linePitch 316 (15.8pt) snaps to 2 rows (31.6). Word
keeps an empty paragraph at 759.8 (box to 791.4) on a page whose body ends at 785.3; Oxi
moves it. Sheet: FILL Meiryo-12 paragraphs ('x' or empty) from the top; the k-th paragraph's
snapped box overhangs by k*31.6 - room. Readout: the page of each of the last 4 paragraphs
(COM Information(3), collapsed start) => how far past the body bottom a box may reach.
Arms: text/empty x top margin {1134, 1200, 1250, 1300, 1350} (shifts the overhang in 3.3pt steps).

    python tools/metrics/_pb_gridbottom_fit_gen.py
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/gridbottom_fit'); OUT.mkdir(parents=True, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
RPR = '<w:rPr><w:rFonts w:ascii="メイリオ" w:eastAsia="メイリオ" w:hAnsi="メイリオ"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'


def document(top, text, n=30, sz=24):
    R = RPR.replace('w:val="24"', f'w:val="{sz}"')
    body = ''.join(f'<w:p><w:pPr><w:widowControl/><w:jc w:val="left"/>{R}</w:pPr>' + (f'<w:r>{R}<w:t>{text}{i}</w:t></w:r>' if text else '') + '</w:p>' for i in range(n))
    sect = (f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="{top}" w:right="1134" w:bottom="1134" w:left="1134" w:header="851" w:footer="284" w:gutter="0"/>'
            '<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="316"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False
try:
    import os
    SZ = int(os.environ.get('PB_SZ', '24')); PITCH = 31.6 if SZ >= 20 else 15.8
    for text in ('', 'x'):
        for top in (1134, 1160, 1190, 1220, 1250, 1280, 1310, 1340):
            at = OUT / f'{"txt" if text else "empty"}_sz{SZ}_top{top}.docx'
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
                z.writestr('word/settings.xml', SETT); z.writestr('word/document.xml', document(top, text, n=60 if SZ < 20 else 30, sz=SZ))
            d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
            try:
                n1 = 0; last_y = None
                for i in range(1, (61 if SZ < 20 else 31)):
                    rg = d.Paragraphs(i).Range; c = d.Range(rg.Start, rg.Start)
                    if c.Information(3) == 1: n1 += 1; last_y = round(c.Information(6), 2)
                body_bottom = 841.9 - 1134 / 20
                print(f'{"txt" if text else "empty"} top={top/20:6.2f} on_p1={n1:2d} last_p1_y={last_y} box_end={round(last_y + PITCH, 2)} overhang={round(last_y + PITCH - body_bottom, 2)}', flush=True)
            finally:
                d.Close(False)
finally:
    app.Quit()
