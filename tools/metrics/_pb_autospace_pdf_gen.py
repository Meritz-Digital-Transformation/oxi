# -*- coding: utf-8 -*-
"""How wide is the gap Word inserts between an East Asian letter and a digit / Latin
letter (autoSpaceDE / autoSpaceDN), read from Word's own PDF?

policies__1411889624a10ff9 p1 (MS 明朝 10.5pt body, Century ascii, compat 15,
balanceSingleByteDoubleByteWidth, docGrid type=lines): the address line
「　　　　　　〒951-8062　新潟県新潟市中央区西堀前通6番町894-1　西堀6番館ビル5F」
fits Word's 425.2pt line with 0.9pt to spare (glyph origins: kanji 10.5, Century
digit 5.844, six CJK|digit gaps of 2.455 = 0.234em) and wraps in Oxi, whose gap is
fs/4 = 2.625 (S546) -- six gaps put the line 0.05pt over. S1175's table (Info(5),
0.75pt quantised) read the gap as 0.253-0.263 of the advance; the PDF says less.

Sheet: one paragraph per arm, jc=left, no indent: ten repeats of 「漢字9」 (a kanji
pair followed by a digit, so each repeat carries two gaps: 字|9 and 9|漢) plus the
same with a Latin letter 「漢字a」. Readout: per-char origins from the PDF; the gap
is (origin of the char after the boundary) - (origin before + its advance), where
the advance of a kanji is measured from a kanji|kanji pair on the same line and the
digit advance from the 「9999999999」 control line.

Arms: fs {8, 9, 10.5, 11, 12, 14} x balance {on, off} x ascii face {Century, Times
New Roman, Arial} x eastAsia {ＭＳ 明朝, ＭＳ ゴシック}. Compat 15 unless --compat 14.

    python tools/metrics/_pb_autospace_pdf_gen.py            # generate + export + read
    python tools/metrics/_pb_autospace_pdf_gen.py --gen-only
"""
import zipfile, sys, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/autospace_pdf'); OUT.mkdir(parents=True, exist_ok=True)
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


def settings(compat, balance):
    bal = '<w:balanceSingleByteDoubleByteWidth/>' if balance else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:compat>{bal}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')


def styles(ascii_face, ea_face, sz):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{ascii_face}" w:eastAsia="{ea_face}" w:hAnsi="{ascii_face}" w:cs="Times New Roman"/>'
            f'<w:kern w:val="2"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="left"/></w:pPr></w:style></w:styles>')


def para(text):
    return f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def document():
    body = para('漢字9' * 10) + para('漢字a' * 10) + para('9' * 20) + para('a' * 20) + para('漢字' * 10) + para('漢字' + '9' * 5 + '漢字' + 'ab' * 3 + '漢字')
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'


FACES = [('Century', 'cen'), ('Times New Roman', 'tnr'), ('Arial', 'ari')]
EAS = [('ＭＳ 明朝', 'min'), ('ＭＳ ゴシック', 'got')]
SIZES = [8, 9, 10.5, 11, 12, 14]
compat = 14 if '--compat' in sys.argv and '14' in sys.argv else 15
ARMS = []
for fs in SIZES:
    for bal in (1, 0):
        for face, ftag in FACES:
            for ea, etag in EAS:
                ARMS.append((fs, bal, face, ftag, ea, etag))


def name(fs, bal, face, ftag, ea, etag):
    return f'c{compat}_fs{str(fs).replace(".", "p")}_bal{bal}_{ftag}_{etag}.docx'


for arm in ARMS:
    fs, bal, face, ftag, ea, etag = arm
    at = OUT / name(*arm)
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR); z.writestr('word/_rels/document.xml.rels', DR)
        z.writestr('word/settings.xml', settings(compat, bal)); z.writestr('word/styles.xml', styles(face, ea, int(fs * 2)))
        z.writestr('word/document.xml', document())
print(f'{len(ARMS)} docx written to {OUT}')
if '--gen-only' in sys.argv:
    sys.exit(0)

import win32com.client
app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
try:
    for arm in ARMS:
        at = OUT / name(*arm); pdf = at.with_suffix('.pdf')
        if pdf.exists():
            continue
        d = app.Documents.Open(str(at.resolve()), ReadOnly=True)
        try:
            d.ExportAsFixedFormat(str(pdf.resolve()), 17)
        finally:
            d.Close(False)
finally:
    app.Quit()

import fitz


def read(pdf):
    # fitz splits a line at font changes; regroup every glyph by its baseline y
    doc = fitz.open(str(pdf)); pg = doc[0]; rows = {}
    for b in pg.get_text('rawdict')['blocks']:
        for l in b.get('lines', []):
            for sp in l['spans']:
                for ch in sp['chars']:
                    if ch['c'].strip():
                        rows.setdefault(round(ch['origin'][1], 0), []).append((ch['c'], ch['origin'][0], sp['font']))
    lines = [(y, sorted(chars, key=lambda c: c[1])) for y, chars in rows.items()]
    lines.sort()
    return lines


def adv(chars, pred_a, pred_b):
    """mean origin step over consecutive pairs (a, b) satisfying the predicates"""
    v = [b[1] - a[1] for a, b in zip(chars, chars[1:]) if pred_a(a[0]) and pred_b(b[0])]
    return (sum(v) / len(v), len(v)) if v else (None, 0)


kanji = lambda c: 0x4e00 <= ord(c) <= 0x9fff
digit = lambda c: c.isdigit()
latin = lambda c: c.isascii() and c.isalpha()
print(f'{"arm":32s} {"kanji":>6s} {"digit":>6s} {"latin":>6s} | {"字|9":>6s} {"9|漢":>6s} {"字|a":>6s} {"a|漢":>6s} | gap/em (9) (a)')
for arm in ARMS:
    fs, bal, face, ftag, ea, etag = arm
    lines = read((OUT / name(*arm)).with_suffix('.pdf'))
    if len(lines) < 6:
        print(name(*arm), 'lines', len(lines)); continue
    l9, la, ld, ll, lk = [c for _, c in lines[:5]]
    k_adv = adv(lk, kanji, kanji)[0]
    d_adv = adv(ld, digit, digit)[0]
    a_adv = adv(ll, latin, latin)[0]
    # boundaries on the 漢字9 line: 字->9 gap = origin(9) - (origin(字) + k_adv); 9->漢 gap = origin(漢) - (origin(9) + d_adv)
    g_k9 = adv(l9, kanji, digit)[0] - k_adv
    g_9k = adv(l9, digit, kanji)[0] - d_adv
    g_ka = adv(la, kanji, latin)[0] - k_adv
    g_ak = adv(la, latin, kanji)[0] - a_adv
    print(f'{name(*arm)[:-5]:32s} {k_adv:6.3f} {d_adv:6.3f} {a_adv:6.3f} | {g_k9:6.3f} {g_9k:6.3f} {g_ka:6.3f} {g_ak:6.3f} | {(g_k9 + g_9k) / 2 / fs:.4f} {(g_ka + g_ak) / 2 / fs:.4f}')
