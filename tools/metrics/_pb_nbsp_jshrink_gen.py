# -*- coding: utf-8 -*-
"""Does Word shrink a NON-BREAKING SPACE on a justified line, and how much does a
compat-14 justified Latin line squeeze? (reference__0096d8c9 p159 / p230)

Sheet: Times New Roman 11, one document per (arm, jc, compat); each paragraph is the
arm's text with a right indent narrowing the line 1 twip per paragraph. Readout: the
narrowest width that keeps line 1 intact (the flip). allow = left flip - justified flip.
Arms: the last unit joined by a normal space vs by U+00A0 (same widths), and the p230
line itself.

    python tools/metrics/_pb_nbsp_jshrink_gen.py
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/nbsp_jshrink'); OUT.mkdir(parents=True, exist_ok=True)
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


def settings(compat):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')


STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-AU" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
SECT = ('<w:sectPr><w:pgSz w:w="11907" w:h="16839"/><w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000" w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/></w:sectPr>')
TEXT_W = 11907 - 2000
BASE = 'substantiation notice, has the meaning given by the Tribunal under'
ARMS = {
    'sp': BASE + ' subsection 60FC(2). tail words here',
    'nbsp': BASE + ' subsection 60FC(2). tail words here',
    'p230': 'contravention of a provision referred to in that paragraph; and more words to wrap',
}
N = int(os.environ.get('NPARA', '800'))


def doc(text, jc, w0_tw):
    ps = []
    for i in range(N):
        right = TEXT_W - (w0_tw - i)
        ps.append(f'<w:p><w:pPr><w:ind w:right="{right}"/><w:jc w:val="{jc}"/></w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(ps)}{SECT}</w:body></w:document>'


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0


def first_line_end(d, i):
    r = d.Paragraphs(i).Range
    # character index where line 2 starts (Information(10) = wdFirstCharacterLineNumber)
    y0 = d.Range(r.Start, r.Start).Information(6)
    for k in range(r.Start, r.End):
        if d.Range(k, k).Information(6) != y0:
            return k - r.Start
    return r.End - r.Start


def flips(arm, jc, compat, w0_tw, needle):
    at = OUT / f'{arm}_{jc}_c{compat}.docx'
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
        z.writestr('word/settings.xml', settings(compat)); z.writestr('word/styles.xml', STYLES); z.writestr('word/document.xml', doc(ARMS[arm], jc, w0_tw))
    d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
    try:
        # binary search the first paragraph whose line 1 no longer contains the needle's end
        text = ARMS[arm]
        end = text.index(needle) + len(needle)
        lo, hi = 0, N - 1
        if first_line_end(d, 1) < end:
            return None
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if first_line_end(d, mid + 1) >= end: lo = mid
            else: hi = mid - 1
        return (w0_tw - lo) / 20.0
    finally:
        d.Close(False)


try:
    for arm, needle in (('sp', '60FC(2).'), ('nbsp', '60FC(2).'), ('p230', 'and')):
        for compat in (14, 15):
            left = flips(arm, 'left', compat, TEXT_W, needle)
            both = flips(arm, 'both', compat, TEXT_W, needle)
            print(f'{arm:5s} c{compat} keep-{needle!r} left={left} both={both} allow={None if None in (left, both) else round(left - both, 2)}', flush=True)
finally:
    app.Quit()
