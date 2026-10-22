"""Собирает PDF с галочками из markdown-файлов плана.

Запуск: python3 plan/make_pdf.py — PDF появятся в plan/pdf/.
Галочки — настоящие поля формы: их можно отмечать в Acrobat, Chrome, Xodo,
Preview и т. п., а на распечатке это пустые квадратики.
"""
import os
import re
import sys

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Flowable, HRFlowable, KeepTogether, Paragraph,
                                SimpleDocTemplate, Spacer, Table, TableStyle)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'pdf')
FONT_DIR = '/usr/share/fonts/truetype/dejavu'

FILES = [
    ('menu-i-zakupki.md', 'Oktoberfest — меню и закупки.pdf'),
    ('konkursy.md', 'Oktoberfest — конкурсы.pdf'),
    ('rulka.md', 'Oktoberfest — рулька.pdf'),
    ('recepty.md', 'Oktoberfest — рецепты.pdf'),
]

# Таблицы под этими заголовками — справочные, без галочек.
NO_CHECK_HEADINGS = {'Классика Баварии и Германии'}

BLUE = colors.HexColor('#1f5fa8')
LIGHT = colors.HexColor('#e8f0fa')
GRID = colors.HexColor('#b8c7da')
QMARK = '<font color="#c0392b"><b>?</b></font>'

pdfmetrics.registerFont(TTFont('DV', os.path.join(FONT_DIR, 'DejaVuSans.ttf')))
pdfmetrics.registerFont(TTFont('DV-B', os.path.join(FONT_DIR, 'DejaVuSans-Bold.ttf')))
# Курсивного DejaVu Sans в системе нет — курсив рисуем прямым начертанием.
pdfmetrics.registerFontFamily('DV', normal='DV', bold='DV-B', italic='DV', boldItalic='DV-B')

BASE = ParagraphStyle('base', fontName='DV', fontSize=9.5, leading=13, alignment=TA_LEFT)
CELL = ParagraphStyle('cell', parent=BASE, fontSize=8.5, leading=11)
CELL_H = ParagraphStyle('cellh', parent=CELL, fontName='DV-B', textColor=BLUE)
H1 = ParagraphStyle('h1', parent=BASE, fontName='DV-B', fontSize=17, leading=21, textColor=BLUE, spaceAfter=6)
H2 = ParagraphStyle('h2', parent=BASE, fontName='DV-B', fontSize=13, leading=17, textColor=BLUE, spaceBefore=10, spaceAfter=4)
H3 = ParagraphStyle('h3', parent=BASE, fontName='DV-B', fontSize=10.5, leading=14, spaceBefore=7, spaceAfter=3)


def inline(text):
    """Markdown-разметка строки → разметка Paragraph."""
    text = text.replace('\\*', '\x00')
    text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    text = re.sub(r'\[([^\]]+)\]\((https?://[^)]+)\)', r'<a href="\2" color="#1f5fa8"><u>\1</u></a>', text)
    text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
    text = re.sub(r'(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])', r'<i>\1</i>', text)
    text = re.sub(r'`([^`]+)`', r'<font name="DV-B">\1</font>', text)
    text = text.replace('\x00', '*')
    text = (text.replace('✅', '✓').replace('❓', QMARK).replace('✏️', '✎')
            .replace('✏', '✎').replace('🍺', '').replace('️', ''))
    return text.strip()


class CheckBox(Flowable):
    """Поле-галочка формы PDF."""
    counter = 0

    def __init__(self, size=10):
        super().__init__()
        self.size = size
        self.width = self.height = size

    def draw(self):
        CheckBox.counter += 1
        self.canv.acroForm.checkbox(
            name='cb%d' % CheckBox.counter, x=0, y=0, size=self.size,
            buttonStyle='check', borderColor=BLUE, fillColor=colors.white,
            textColor=BLUE, borderWidth=1, forceBorder=True, relative=True)


def check_item(text, avail):
    t = Table([[CheckBox(), Paragraph(inline(text), BASE)]], colWidths=[7 * mm, avail - 7 * mm])
    t.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (0, 0), 2.5), ('TOPPADDING', (1, 0), (1, 0), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5),
    ]))
    return t


def plain_len(s):
    return len(re.sub(r'\*\*|`|\[|\]\([^)]*\)', '', s))


def make_table(rows, avail, with_check):
    ncol = len(rows[0])
    rows = [r + [''] * (ncol - len(r)) for r in rows]
    # Ширины колонок — по длине текста, с ограничениями.
    weights = []
    for c in range(ncol):
        m = max(plain_len(r[c]) for r in rows)
        weights.append(min(max(m, 6), 60))
    cw = 7 * mm if with_check else 0
    rest = avail - cw
    widths = [rest * w / sum(weights) for w in weights]
    data = []
    for i, r in enumerate(rows):
        style = CELL_H if i == 0 else CELL
        cells = [Paragraph(inline(x), style) for x in r]
        if with_check:
            cells = ([''] if i == 0 else [CheckBox(9)]) + cells
        data.append(cells)
    t = Table(data, colWidths=([cw] if with_check else []) + widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BACKGROUND', (0, 0), (-1, 0), LIGHT),
        ('LINEBELOW', (0, 0), (-1, -1), 0.4, GRID),
        ('BOX', (0, 0), (-1, -1), 0.6, GRID),
        ('LEFTPADDING', (0, 0), (-1, -1), 4), ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 3), ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    return t


def parse(md, avail):
    story, para, heading = [], [], ''
    lines = md.split('\n')
    i = 0

    def flush():
        if para:
            story.append(Paragraph(inline(' '.join(para)), BASE))
            story.append(Spacer(1, 4))
            para.clear()

    while i < len(lines):
        line = lines[i].rstrip()
        s = line.strip()
        if not s:
            flush(); i += 1; continue
        if s.startswith('|'):
            flush()
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
                if not all(re.fullmatch(r':?-{2,}:?', c) for c in cells):
                    rows.append(cells)
                i += 1
            story.append(make_table(rows, avail, heading not in NO_CHECK_HEADINGS))
            story.append(Spacer(1, 6))
            continue
        m = re.match(r'(#{1,3})\s+(.*)', s)
        if m:
            flush()
            level, text = len(m.group(1)), m.group(2)
            heading = re.sub(r'[*`]', '', text).strip()
            style = {1: H1, 2: H2, 3: H3}[level]
            story.append(Paragraph(inline(text), style))
            i += 1; continue
        if s == '---':
            flush()
            story.append(HRFlowable(width='100%', thickness=0.6, color=GRID, spaceBefore=4, spaceAfter=4))
            i += 1; continue
        m = re.match(r'- \[[ xX]\]\s+(.*)', s)
        if m:
            flush(); story.append(check_item(m.group(1), avail)); i += 1; continue
        m = re.match(r'(\d+)\.\s+(.*)', s)
        if m:
            flush()
            story.append(Paragraph(inline(m.group(2)), BASE, bulletText=m.group(1) + '.'))
            i += 1; continue
        m = re.match(r'[-*]\s+(.*)', s)
        if m:
            flush()
            story.append(Paragraph(inline(m.group(1)), BASE, bulletText='•'))
            i += 1; continue
        para.append(s)
        i += 1
    flush()
    return story


def footer(canv, doc):
    canv.saveState()
    canv.setFont('DV', 7.5)
    canv.setFillColor(colors.grey)
    canv.drawString(15 * mm, 9 * mm, doc.title)
    canv.drawRightString(A4[0] - 15 * mm, 9 * mm, 'стр. %d' % doc.page)
    canv.restoreState()


def build(src, dst):
    md = open(os.path.join(HERE, src), encoding='utf-8').read()
    title = re.search(r'^#\s+(.*)$', md, re.M).group(1)
    doc = SimpleDocTemplate(os.path.join(OUT, dst), pagesize=A4,
                            leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=14 * mm, bottomMargin=16 * mm,
                            title=title, author='Oktoberfest 24.10')
    # Списки с bulletText требуют отступа.
    for st in (BASE,):
        st.bulletFontName = 'DV'
    story = []
    for f in parse(md, doc.width):
        if isinstance(f, Paragraph) and f.bulletText:
            f.style = ParagraphStyle('li', parent=BASE, leftIndent=12, bulletIndent=2, spaceAfter=1.5)
        story.append(f)
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return os.path.join(OUT, dst)


if __name__ == '__main__':
    os.makedirs(OUT, exist_ok=True)
    for src, dst in FILES:
        print(build(src, dst))
    sys.exit(0)
