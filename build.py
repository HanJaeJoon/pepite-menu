# -*- coding: utf-8 -*-
"""
페피트 메뉴판 빌더
------------------
data/*.json 의 내용으로 docx 메뉴판을 생성한다.

레이아웃 원칙
  - 메뉴 1개 = 표 1행 (왼쪽 셀: 이름/설명, 오른쪽 셀: 가격)
  - 표 테두리는 전부 없음 -> 화면상 보이지 않지만 위치는 고정
  - 행 분할 금지(cantSplit) -> 페이지 경계에서 한 메뉴가 잘리지 않음
  - 섹션 헤더는 keepNext -> 헤더만 페이지 끝에 남지 않음
  - 페이지 구성은 PAGES 에 명시 -> 앞 페이지 내용이 늘어도 뒤 페이지가 안 밀림

수정 방법
  - 메뉴 내용: data/*.json 편집
  - 글자 크기/여백/페이지 구성: 아래 STYLE, PAGES 편집
  - 실행:  python build.py           (docx 생성)
          python build.py --debug   (표 테두리를 보이게 해서 배치 확인)
"""
import json, os, sys

import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, 'data')
DIST = os.path.join(HERE, 'dist')
os.makedirs(DIST, exist_ok=True)
DEBUG = '--debug' in sys.argv

# --------------------------------------------------------------------------
# 스타일 (여기만 고치면 전체 메뉴판에 반영됨)
# --------------------------------------------------------------------------
# 폰트는 원본 4개 파일이 문서마다 달랐다.
#   요리/런치/코스 = 바탕 + Cambria(테마 기본), 와인 = 맑은 고딕 + Times New Roman,
#   코스 영문에는 서울남산체 L 이 20군데 섞여 있었음.
# 다수(요리/런치/코스)의 명조 기조를 기준으로 전 문서를 통일했다.
STYLE = {
    'font_head': 'ChunkFive Ex',    # 로고 / 섹션 헤더 / 가격
    'font_kor': '바탕',              # 한글 본문 (원본 요리/런치/코스와 동일)
    'font_eng': 'Times New Roman',  # 영문 본문 (원본 와인과 동일, Cambria 와 같은 명조 계열)

    'title_size': 22,              # PÉPITE
    'section_size': 18,            # STARTER, PASTA ...
    'name_size': 11.5,             # 메뉴 한글명
    'engname_size': 9,             # 메뉴 영문명
    'desc_size': 9,                # 설명
    'price_size': 11.5,

    # 브랜드 색 (기존 메뉴판에서 그대로 가져옴)
    'logo': RGBColor(0xE3, 0x6C, 0x0A),    # PÉPITE 로고 주황
    'gray': RGBColor(0x40, 0x40, 0x40),    # 영문 보조 텍스트
    'header': RGBColor(0x00, 0x00, 0x00),  # 섹션 헤더는 검정

    # 제목 아래 가로선 (원본과 동일: single 0.5pt, 글자와 간격 14)
    'rule_under_title': True,
    'rule_under_section': False,   # 원본에는 없음. True 로 켜면 섹션 헤더 아래에도 선
    'rule_size': 4,                # 1/8 pt 단위 -> 4 = 0.5pt
    'rule_space': 14,

    'item_gap': 9,                 # 메뉴 사이 간격(pt)
    'section_gap_before': 16,      # 섹션 헤더 위 간격
    'section_gap_after': 8,

    'page_w': 21.0, 'page_h': 29.7,          # A4 (cm)
    'margin_lr': 2.2, 'margin_tb': 1.6,
    'price_col': 2.4,                        # 가격 칸 폭(cm)
}

# 와인은 항목 수가 많아 본문보다 한 단계 작게 (원본 와인 메뉴판과 비슷한 밀도)
WINE_STYLE = {
    'name_size': 9,
    'engname_size': 8,
    'region_size': 7,
    'desc_size': 8,
    'price_size': 9.5,
    'item_gap': 6,
}

# 페이지 구성: 리스트 하나 = 한 페이지에 들어갈 섹션들
PAGES = {
    'food': [['STARTER'], ['PASTA'], ['MAIN', 'MEAL'], ['SIDE', 'DESSERT']],
    'lunch': [['STARTER'], ['PASTA', 'MAIN'], ['MEAL', 'DESSERT', 'SIDE', 'DRINKS']],
    'wine': [['WHITE WINE'], ['RED WINE'], ['SPARKLING WINE', 'HOUSE WINE']],
}


# --------------------------------------------------------------------------
# 저수준 헬퍼
# --------------------------------------------------------------------------
def set_cell_borders_none(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + edge)
        e.set(qn('w:val'), 'single' if DEBUG else 'none')
        e.set(qn('w:sz'), '4')
        e.set(qn('w:color'), 'FF0000' if DEBUG else 'auto')
        borders.append(e)
    tblPr.append(borders)


def set_cell_margins(table, top=0, bottom=0, left=0, right=0):
    tblPr = table._tbl.tblPr
    mar = OxmlElement('w:tblCellMar')
    for side, val in (('top', top), ('left', left), ('bottom', bottom), ('right', right)):
        e = OxmlElement('w:' + side)
        e.set(qn('w:w'), str(int(val * 567)))   # cm -> twips
        e.set(qn('w:type'), 'dxa')
        mar.append(e)
    tblPr.append(mar)


def no_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:cantSplit'))


def bottom_rule(p, color=None):
    """문단 아래 가로선(hr). 원본 메뉴판의 PÉPITE 밑줄과 동일한 속성."""
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single')
    b.set(qn('w:sz'), str(STYLE['rule_size']))
    b.set(qn('w:space'), str(STYLE['rule_space']))
    b.set(qn('w:color'), '%02X%02X%02X' % tuple(color or STYLE['logo']))
    bdr.append(b)
    pPr.append(bdr)
    return p


def keep_next(p, on=True):
    pPr = p._p.get_or_add_pPr()
    e = OxmlElement('w:keepNext')
    if not on:
        e.set(qn('w:val'), '0')
    pPr.append(e)
    pPr.append(OxmlElement('w:keepLines'))


def line(cell, text, *, font, size, bold=False, color=None,
         after=0, first=False, align=None):
    """셀 안에 한 줄(문단) 추가. first=True 면 셀의 기본 빈 문단을 재사용."""
    p = cell.paragraphs[0] if first else cell.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(after)
    pf.line_spacing = 1.0
    if align is not None:
        p.alignment = align
    r = p.add_run(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    if color is not None:
        r.font.color.rgb = color
    r._element.rPr.rFonts.set(qn('w:eastAsia'), STYLE['font_kor'])
    return p


def section_header(doc, name):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(STYLE['section_gap_before'])
    pf.space_after = Pt(STYLE['section_gap_after'])
    keep_next(p)
    r = p.add_run(name)
    r.font.name = STYLE['font_head']
    r.font.size = Pt(STYLE['section_size'])
    r.font.color.rgb = STYLE['header']
    r._element.rPr.rFonts.set(qn('w:eastAsia'), STYLE['font_head'])
    if STYLE['rule_under_section']:
        bottom_rule(p)
    return p


def doc_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    keep_next(p)
    r = p.add_run(text)
    r.font.name = STYLE['font_head']
    r.font.size = Pt(STYLE['title_size'])
    r.font.color.rgb = STYLE['logo']       # 브랜드 주황
    r._element.rPr.rFonts.set(qn('w:eastAsia'), STYLE['font_head'])
    if STYLE['rule_under_title']:
        bottom_rule(p)
    return p


def new_doc():
    doc = docx.Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(STYLE['page_w']), Cm(STYLE['page_h'])
    s.left_margin = s.right_margin = Cm(STYLE['margin_lr'])
    s.top_margin = s.bottom_margin = Cm(STYLE['margin_tb'])
    n = doc.styles['Normal']
    n.font.name = STYLE['font_kor']
    n.font.size = Pt(STYLE['desc_size'])
    n.element.rPr.rFonts.set(qn('w:eastAsia'), STYLE['font_kor'])
    return doc


def item_table(doc, n_rows):
    body_w = STYLE['page_w'] - 2 * STYLE['margin_lr']
    t = doc.add_table(rows=n_rows, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    set_cell_borders_none(t)
    set_cell_margins(t)
    for row in t.rows:
        no_split(row)
        row.cells[0].width = Cm(body_w - STYLE['price_col'])
        row.cells[1].width = Cm(STYLE['price_col'])
    return t


def price_cell(cell, price, size=None):
    line(cell, price, font=STYLE['font_head'], size=size or STYLE['price_size'],
         first=True, align=WD_ALIGN_PARAGRAPH.RIGHT)


# --------------------------------------------------------------------------
# 메뉴판별 조립
# --------------------------------------------------------------------------
def build_dish(key, out_name):
    data = json.load(open(os.path.join(DATA, key + '.json'), encoding='utf-8'))
    by_name = {s['name']: s for s in data['sections']}
    doc = new_doc()

    for pi, page in enumerate(PAGES[key]):
        if pi:
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        doc_title(doc, data['title'])
        for sec in page:
            s = by_name[sec]
            section_header(doc, s['name'])
            t = item_table(doc, len(s['items']))
            for row, it in zip(t.rows, s['items']):
                c = row.cells[0]
                line(c, it['kor'], font=STYLE['font_kor'], size=STYLE['name_size'],
                     bold=True, first=True, after=1)
                if it['eng']:
                    line(c, it['eng'], font=STYLE['font_eng'],
                         size=STYLE['engname_size'], bold=True, after=2)
                if it['desc_kor']:
                    line(c, '* ' + it['desc_kor'], font=STYLE['font_kor'],
                         size=STYLE['desc_size'], after=1)
                if it['desc_eng']:
                    line(c, it['desc_eng'], font=STYLE['font_eng'],
                         size=STYLE['desc_size'], color=STYLE['gray'])
                c.paragraphs[-1].paragraph_format.space_after = Pt(STYLE['item_gap'])
                price_cell(row.cells[1], it['price'])

    doc.save(os.path.join(DIST, out_name))
    print('saved', out_name)


def build_wine(out_name):
    data = json.load(open(os.path.join(DATA, 'wine.json'), encoding='utf-8'))
    by_name = {s['name']: s for s in data['sections']}
    doc = new_doc()

    for pi, page in enumerate(PAGES['wine']):
        if pi:
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        for sec in page:
            s = by_name[sec]
            section_header(doc, s['name'])
            t = item_table(doc, len(s['items']))
            for row, it in zip(t.rows, s['items']):
                c = row.cells[0]
                line(c, it['kor'], font=STYLE['font_kor'], size=WINE_STYLE['name_size'],
                     bold=True, first=True, after=0)
                if it['eng']:
                    line(c, it['eng'], font=STYLE['font_eng'],
                         size=WINE_STYLE['engname_size'], bold=True, after=1)
                if it.get('region'):
                    line(c, it['region'], font=STYLE['font_kor'],
                         size=WINE_STYLE['region_size'], color=STYLE['gray'], after=1)
                if it['desc_kor']:
                    line(c, it['desc_kor'], font=STYLE['font_kor'],
                         size=WINE_STYLE['desc_size'], after=0)
                if it['desc_eng']:
                    line(c, it['desc_eng'], font=STYLE['font_eng'],
                         size=WINE_STYLE['desc_size'], color=STYLE['gray'])
                c.paragraphs[-1].paragraph_format.space_after = Pt(WINE_STYLE['item_gap'])
                price_cell(row.cells[1], it['price'], WINE_STYLE['price_size'])

    doc.save(os.path.join(DIST, out_name))
    print('saved', out_name)


def build_course(out_name):
    data = json.load(open(os.path.join(DATA, 'course.json'), encoding='utf-8'))
    doc = docx.Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(14.8), Cm(21.0)     # A5
    s.left_margin = s.right_margin = Cm(1.6)
    s.top_margin = s.bottom_margin = Cm(1.4)
    n = doc.styles['Normal']
    n.font.name = STYLE['font_kor']
    n.element.rPr.rFonts.set(qn('w:eastAsia'), STYLE['font_kor'])

    heads = {'STARTER', 'PASTA', 'MAIN', 'DESSERT'}

    def put(text, size, *, font=None, bold=False, before=0, after=0, color=None):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.line_spacing = 1.0
        keep_next(p)
        r = p.add_run(text)
        r.font.name = font or STYLE['font_kor']
        r.font.size = Pt(size)
        r.font.bold = bold
        if color is not None:
            r.font.color.rgb = color
        # 한글 대체 폰트는 헤드 폰트를 쓰는 줄(로고/섹션명)만 예외로 둔다
        ea = font if font == STYLE['font_head'] else STYLE['font_kor']
        r._element.rPr.rFonts.set(qn('w:eastAsia'), ea)
        return p

    t = put(data['title'], 20, font=STYLE['font_head'], after=2, color=STYLE['logo'])
    if STYLE['rule_under_title']:
        bottom_rule(t)
    put(data['subtitle'], 12, font=STYLE['font_head'], after=6, color=STYLE['logo'])

    for ln in data['lines']:
        if ln in heads:
            put(ln, 12, font=STYLE['font_head'], before=12, after=6)
        elif ln == 'or':
            put('or', 8, color=STYLE['gray'], before=1, after=1)
        elif ln.startswith('98,000'):
            put(ln, 16, font=STYLE['font_head'], before=20, after=2)
        elif ln.startswith('(1인기준'):
            put(ln, 9, color=STYLE['gray'])
        elif ln.startswith('(') or ln.startswith('Pre-meal') or ln.startswith('(Add'):
            put(ln, 6, color=STYLE['gray'], after=1)
        elif ln[0].isascii():
            put(ln, 8, font=STYLE['font_eng'], color=STYLE['gray'], after=3)
        else:
            put(ln, 10, bold=True, after=1)

    doc.save(os.path.join(DIST, out_name))
    print('saved', out_name)


if __name__ == '__main__':
    build_dish('food', '페피트 요리 메뉴판.docx')
    build_dish('lunch', '페피트 런치 메뉴판.docx')
    build_wine('페피트 와인 메뉴판.docx')
    build_course('페피트 런치 코스메뉴.docx')
