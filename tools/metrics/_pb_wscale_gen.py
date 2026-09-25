# -*- coding: utf-8 -*-
"""How wide is a `w:w`-scaled CJK run on a linesAndChars character grid?
reference__530de8d7 (JA blind-F, docGrid linesAndChars linePitch 411 charSpace
3042 -> char pitch 11.74pt at 11pt): a 2-column table cell holds
'・母子家庭及び父子家庭高等職業訓練促進給付金等事業（155㌻）' (30 chars) in
ＭＳ Ｐ明朝 with <w:w w:val="66"/>.  Word keeps it on ONE line (row pitch stays
one line); Oxi breaks after 28 chars.  Oxi's per-char advance is 8.00 =
11 x 0.66 + 0.74 -- the fullwidth grid increment (pitch - em) is added AFTER
the scale.  Word must be at <= 7.6/char for 30 chars to fit 229.75pt, i.e.
either 11 x 0.66 = 7.26 (no grid increment on a scaled run) or 11.74 x 0.66 =
7.75 (increment scaled too; 30 x 7.75 = 232.5 does NOT fit, so unlikely).
Sheet: A4, Normal = ＭＳ 明朝 11pt (docDefaults eastAsia).  One body paragraph
per arm holding 30 kanji (no punctuation), jc=left, so every glyph x in the PDF
is a pure advance.  Arms: grid {lc 411/3042, lines 411, none} x w {100, 66, 80}
x eastAsia font {ＭＳ 明朝, ＭＳ Ｐ明朝}.  Readout: PDF char origins (fitz
rawdict) -> per-char advance; Oxi --dump-layout per-char x -> the same.
Usage: python tools/metrics/_pb_wscale_gen.py   (writes tests/fixtures/wscale/*.docx)
"""
import zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = Path('tests/fixtures/wscale'); OUT.mkdir(parents=True, exist_ok=True)
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
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ' + W + '>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + W + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>'
          '<w:kern w:val="2"/><w:sz w:val="22"/><w:szCs w:val="22"/>'
          '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>')
GRIDS = {'lc': '<w:docGrid w:type="linesAndChars" w:linePitch="411" w:charSpace="3042"/>',
         'lines': '<w:docGrid w:type="lines" w:linePitch="411"/>',
         'none': ''}
KANJI = ('母子家庭及父子家庭高等職業訓練促進給付金等事業' * 2)[:30]  # 30 chars, no punctuation
assert len(KANJI) == 30
FONTS = {'mincho': 'ＭＳ 明朝', 'pmincho': 'ＭＳ Ｐ明朝'}


def para(text, font, wval):
    rpr = '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>' % (font, font, font)
    if wval != 100:
        rpr += '<w:w w:val="%d"/>' % wval
    return ('<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (rpr, text))


def cell_table(text, font, wval):
    # the real case: a 2-column table (gridCol 5046 x 2), the run in the RIGHT cell,
    # paragraph ind leftChars=100 left=235, line=280 exact -- as reference__530de8d7
    tc = ('<w:tc><w:tcPr><w:tcW w:w="5046" w:type="dxa"/></w:tcPr>'
          '<w:p><w:pPr><w:spacing w:line="280" w:lineRule="exact"/><w:ind w:leftChars="100" w:left="235"/></w:pPr>%s</w:p></w:tc>')
    def run(t, w):
        rpr = '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>' % (font, font, font)
        if w != 100:
            rpr += '<w:w w:val="%d"/>' % w
        return '<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, t)
    return ('<w:tbl><w:tblPr><w:tblW w:w="10092" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5046"/><w:gridCol w:w="5046"/></w:tblGrid>'
            '<w:tr>' + tc % run('左', 100) + tc % run(text, wval) + '</w:tr></w:tbl>')


def doc(grid, font, wval, cell=False, long=False):
    text = (KANJI * 2) if long else KANJI
    body = (cell_table(KANJI, font, wval) if cell else para(text, font, wval)) + para('END', font, 100)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1021" w:bottom="1134" w:left="1021" w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:cols w:space="425"/>%s</w:sectPr>' % GRIDS[grid])
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + W + '><w:body>' + body + sect + '</w:body></w:document>')


def write(path, xml):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
        z.writestr('word/_rels/document.xml.rels', DR); z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/styles.xml', STYLES); z.writestr('word/document.xml', xml)


def cell4000(text, font, wval, ind):
    # non-marginal cell: tcW 4000 (200pt), explicit tblCellMar 108/108, optional leftChars=100
    ppr = '<w:pPr><w:spacing w:line="280" w:lineRule="exact"/>%s</w:pPr>' % ('<w:ind w:leftChars="100" w:left="235"/>' if ind else '')
    def run(t, w):
        rpr = '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:hint="eastAsia"/>' % (font, font, font)
        if w != 100:
            rpr += '<w:w w:val="%d"/>' % w
        return '<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, t)
    tc = '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p>' + ppr + '%s</w:p></w:tc>'
    return ('<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
            '<w:tr>' + tc % run('左', 100) + tc % run(text, wval) + '</w:tr></w:tbl>')


def doc4000(grid, font, wval, ind):
    body = cell4000(KANJI * 2, font, wval, ind) + para('END', font, 100)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1021" w:bottom="1134" w:left="1021" w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:cols w:space="425"/>%s</w:sectPr>' % GRIDS[grid])
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + W + '><w:body>' + body + sect + '</w:body></w:document>')


GLYPH_ARMS = {
    # per-glyph-class advances of a w:w=66 ＭＳ Ｐ明朝 run on the lc grid (Word PDF
    # reference__530de8d7: kanji 7.91/8.03 alternating, び 7.55, （ 4.56, digits
    # 3.96, ㌻ 7.68, ） 4.20; the unscaled ＭＳ 明朝 ・ before it 11.28)
    'kana': 'びびびびびびびびびびびびびびびびびびびび',
    'paren': '（１２３）（１２３）（１２３）（１２３）（１２３）（１２３）',
    'digits': '155155155155155155155155155155',
    'sqm': '㌻㌻㌻㌻㌻㌻㌻㌻㌻㌻㌻㌻㌻㌻㌻㌻㌻㌻㌻㌻',
    'mixed': '母子家庭及び父子家庭高等職業訓練促進給付金等事業（155㌻）',
}


def glyph_doc(grid, key, wval, lead_dot):
    runs = ''
    if lead_dot:
        runs += '<w:r><w:t>・</w:t></w:r>'
    rpr = '<w:rFonts w:ascii="ＭＳ Ｐ明朝" w:eastAsia="ＭＳ Ｐ明朝"/>' + ('<w:w w:val="%d"/>' % wval if wval != 100 else '')
    runs += '<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, GLYPH_ARMS[key])
    body = '<w:p><w:pPr><w:jc w:val="left"/></w:pPr>%s</w:p>' % runs + para('END', 'ＭＳ 明朝', 100)
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1021" w:bottom="1134" w:left="1021" w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:cols w:space="425"/>%s</w:sectPr>' % GRIDS[grid])
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + W + '><w:body>' + body + sect + '</w:body></w:document>')


if __name__ == '__main__':
    n = 0
    for key in GLYPH_ARMS:
        for wval in (100, 66):
            for grid in ('lc', 'none'):
                write(OUT / f'g_{key}_{grid}_w{wval}.docx', glyph_doc(grid, key, wval, key == 'mixed')); n += 1
    for wval in (100, 66, 80):
        for ind in (True, False):
            p = OUT / f'c4k_lc_pmincho_w{wval}_{"ind" if ind else "noind"}.docx'
            write(p, doc4000('lc', FONTS['pmincho'], wval, ind)); n += 1
    for grid in GRIDS:
        for fk, font in FONTS.items():
            for wval in (100, 66, 80):
                p = OUT / f'{grid}_{fk}_w{wval}.docx'
                write(p, doc(grid, font, wval)); n += 1
                if fk == 'pmincho' and wval != 80:
                    p = OUT / f'cell_{grid}_{fk}_w{wval}.docx'
                    write(p, doc(grid, font, wval, cell=True)); n += 1
                    # long body line (60 kanji, jc=left): does the BODY break by em or by pitch?
                    p = OUT / f'long_{grid}_{fk}_w{wval}.docx'
                    write(p, doc(grid, font, wval, long=True)); n += 1
    print('wrote', n, 'arms to', OUT)
