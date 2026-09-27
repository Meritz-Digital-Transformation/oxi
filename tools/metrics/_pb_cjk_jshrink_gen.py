# -*- coding: utf-8 -*-
"""How much does Word shrink a JUSTIFIED CJK line with Latin islands to keep it on one line?

policies__1411889624 (MS Mincho 10.5 + Century, balanceSingleByteDoubleByteWidth, compat 15,
docGrid lines 365): a 48-char address line whose natural width (fs/4 autospace gaps) is
~426.7 stays on one line down to ~421.9 when justified (allow 4.6-5.2); left-aligned it
wraps at its natural width. The Latin law (justify_shrink_two_ceilings) is
allow = min(0.25 x sum(spaces), 0.35 x (last word + one space)).

Sheet: one document per (arm, jc); each paragraph is the arm's text with a right indent
that shrinks the line 1 twip per paragraph from W0 down. Readout: the first paragraph (from
the widest) that takes 2 lines -> min width that keeps 1 line. allow = left flip - both flip.

    python tools/metrics/_pb_cjk_jshrink_gen.py
"""
import zipfile, sys, re, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/cjk_jshrink'); OUT.mkdir(parents=True, exist_ok=True)
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


def settings(balance):
    b = '<w:balanceSingleByteDoubleByteWidth/>' if balance else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:characterSpacingControl w:val="compressPunctuation"/><w:compat>{b}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
          '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/></w:pPr></w:style></w:styles>')
# page: left/right margin 1134 each -> text width 9638tw = 481.9pt; the right indent narrows it
SECT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="365"/></w:sectPr>')
TEXT_W = 9638
ARMS = {
    # the real line (6 gaps, last token 5F)
    'real': '　　　　　　〒951-8062　新潟県新潟市中央区西堀前通6番町894-1　西堀6番館ビル5F',
    # same, long Latin last word (last-word ceiling far away)
    'real_long': '　　　　　　〒951-8062　新潟県新潟市中央区西堀前通6番町894-1　西堀6番館ビルABCDEFGHIJKL',
    # same, CJK last character (no Latin tail)
    'real_cjk_end': '　　　　　　〒951-8062　新潟県新潟市中央区西堀前通6番町894-1　西堀6番館ビル五階',
    # one Latin island mid-line (2 gaps), CJK ending
    'g2_mid': '新潟県新潟市中央区西堀前通新潟市中央区西堀前通ABC新潟県新潟市中央区西堀前通新潟市',
    # one Latin island at the end (1 gap)
    'g1_end': '新潟県新潟市中央区西堀前通新潟市中央区西堀前通新潟県新潟市中央区西堀前通新潟市AB',
    # k mid-line islands 'AB' (2k gaps), CJK ending, total ~41 cells
    'k1': '新潟県新潟市中央区西堀前通新潟市AB中央区西堀前通新潟県新潟市中央区西堀前通新潟',
    'k2': '新潟県新潟市中央区AB西堀前通新潟市中央区西堀AB前通新潟県新潟市中央区西堀前通',
    'k3': '新潟県新潟AB市中央区西堀前通新潟AB市中央区西堀前通新潟県AB新潟市中央区西堀前',
    'k4': '新潟県AB新潟市中央区西堀AB前通新潟市中央AB区西堀前通新潟県新AB潟市中央区西',
    # k islands of ONE letter
    'k2s': '新潟県新潟市中央区A西堀前通新潟市中央区西堀B前通新潟県新潟市中央区西堀前通新潟',
    'rt_5': '　　　　　　〒951-8062　新潟県新潟市中央区西堀前通6番町894-1　西堀6番館ビル5',
    'rt_5FG': '　　　　　　〒951-8062　新潟県新潟市中央区西堀前通6番町894-1　西堀6番館ビル5FG',
    'rt_5FGH': '　　　　　　〒951-8062　新潟県新潟市中央区西堀前通6番町894-1　西堀6番館ビル5FGH',
    'rt_ABCDEF': '　　　　　　〒951-8062　新潟県新潟市中央区西堀前通6番町894-1　西堀6番館ビルABCDEF',
    'ns_5F': '新潟県新潟市〒951-8062新潟県新潟市中央区西堀前通6番町894-1西堀6番館ビル5F',
    'ns_五階': '新潟県新潟市〒951-8062新潟県新潟市中央区西堀前通6番町894-1西堀6番館ビル五階',
    'g1_A': '新潟県新潟市中央区西堀前通新潟市中央区西堀前通新潟県新潟市中央区西堀前通新潟A',
    'g1_ABCD': '新潟県新潟市中央区西堀前通新潟市中央区西堀前通新潟県新潟市中央区西堀前通新潟ABCD',
    'g1_ABCDEF': '新潟県新潟市中央区西堀前通新潟市中央区西堀前通新潟県新潟市中央区西堀前通新潟ABCDEF',
    # pure CJK
    'cjk': '新潟県新潟市中央区西堀前通新潟市中央区西堀前通新潟県新潟市中央区西堀前通新潟市中央区',
}
N = int(os.environ.get('NPARA', '400'))   # 400 twips = 20pt sweep


def runs(text):
    out = []
    for m in re.finditer(r'[0-9A-Za-z-]+|[^0-9A-Za-z-]+', text):
        t = m.group(0)
        hint = '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>' if (t[0].isascii() and t[0].isalnum()) else ''
        out.append(f'<w:r>{hint}<w:t xml:space="preserve">{t}</w:t></w:r>')
    return ''.join(out)


def doc(text, jc, w0_tw):
    ps = []
    for i in range(N):
        right = TEXT_W - (w0_tw - i)
        ps.append(f'<w:p><w:pPr><w:ind w:right="{right}"/><w:jc w:val="{jc}"/></w:pPr>{runs(text)}</w:p>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(ps)}{SECT}</w:body></w:document>'


def build(arm, jc, balance, w0_tw):
    at = OUT / f'{arm}_{jc}_bal{int(balance)}.docx'
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
        z.writestr('word/settings.xml', settings(balance)); z.writestr('word/styles.xml', STYLES); z.writestr('word/document.xml', doc(ARMS[arm], jc, w0_tw))
    return at


import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0


def flip(arm, jc, balance, w0_tw):
    at = build(arm, jc, balance, w0_tw)
    d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
    try:
        for i in range(N):
            if d.Paragraphs(i + 1).Range.ComputeStatistics(1) > 1:
                return (w0_tw - i + 1) / 20.0   # the last width that held one line
        return None
    finally:
        d.Close(False)


try:
    arms = sys.argv[1:] or list(ARMS)
    for arm in arms:
        for balance in ((True, False) if os.environ.get('BAL_BOTH') else (True,)):
            # coarse: find the left flip from a wide start
            left = flip(arm, 'left', balance, TEXT_W)
            if left is None:
                print(arm, 'left never wraps in window'); continue
            w0 = int(round(left * 20)) + 20
            both = flip(arm, 'both', balance, w0)
            print(f'{arm:13s} bal={int(balance)} natural(left)={left:7.2f} justified min={both if both is None else round(both,2)} allow={None if both is None else round(left - both, 2)}', flush=True)
finally:
    app.Quit()
