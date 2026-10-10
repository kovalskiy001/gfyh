# -*- coding: utf-8 -*-
"""Рушій верстки PDF-посібника: шрифти, стилі, блоки, діаграми, шаблони сторінок."""
import re
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor, white
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer,
                                PageBreak, Table, TableStyle, CondPageBreak, KeepTogether)
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.graphics.shapes import Drawing, Rect, String, Polygon, Line, Circle

# ---------- Шрифти ----------
FD = '/usr/share/fonts/truetype/crosextra/'
pdfmetrics.registerFont(TTFont('Carlito', FD + 'Carlito-Regular.ttf'))
pdfmetrics.registerFont(TTFont('Carlito-Bold', FD + 'Carlito-Bold.ttf'))
pdfmetrics.registerFont(TTFont('Carlito-Italic', FD + 'Carlito-Italic.ttf'))
pdfmetrics.registerFont(TTFont('Carlito-BoldItalic', FD + 'Carlito-BoldItalic.ttf'))
addMapping('Carlito', 0, 0, 'Carlito')
addMapping('Carlito', 1, 0, 'Carlito-Bold')
addMapping('Carlito', 0, 1, 'Carlito-Italic')
addMapping('Carlito', 1, 1, 'Carlito-BoldItalic')
pdfmetrics.registerFont(TTFont('Sym', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('JP', '/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf'))

# ---------- Кольори ----------
PRIMARY = HexColor('#1E1B4B')
PRIMARY2 = HexColor('#312E81')
ACCENT = HexColor('#F97316')
TEXT = HexColor('#1F2937')
MUTED = HexColor('#6B7280')
LINE = HexColor('#E5E7EB')
ZEBRA = HexColor('#F8FAFC')
SOFT = HexColor('#C7D2FE')

BOX = {
    'example': ('ПРИКЛАД', '#2563EB', '#EFF6FF'),
    'pro': ('ПОРАДА ПРОФІ', '#059669', '#ECFDF5'),
    'mistake': ('ТИПОВА ПОМИЛКА', '#DC2626', '#FEF2F2'),
    'task': ('ПРАКТИКА', '#D97706', '#FFFBEB'),
    'formula': ('ФОРМУЛА', '#7C3AED', '#F5F3FF'),
    'key': ('ГОЛОВНЕ', '#312E81', '#EEF2FF'),
    'warn': ('ВАЖЛИВО', '#B45309', '#FFF7ED'),
    'quote': ('ЦИТАТА', '#F97316', '#FFF7ED'),
}

PAGE_W, PAGE_H = A4
LM = RM = 52
TM, BM = 62, 56
CW = PAGE_W - LM - RM  # ширина контенту

# ---------- Стилі ----------
S = {}
S['body'] = ParagraphStyle('body', fontName='Carlito', fontSize=10.8, leading=15.2,
                           textColor=TEXT, spaceAfter=6.5, alignment=TA_LEFT)
S['lead'] = ParagraphStyle('lead', parent=S['body'], fontSize=12.2, leading=17,
                           textColor=PRIMARY2, spaceAfter=9)
S['h1plain'] = ParagraphStyle('h1plain', fontName='Carlito-Bold', fontSize=24, leading=29,
                              textColor=PRIMARY, spaceAfter=12)
S['h2'] = ParagraphStyle('h2', fontName='Carlito-Bold', fontSize=16, leading=20,
                         textColor=PRIMARY, spaceBefore=12, spaceAfter=7)
S['h3'] = ParagraphStyle('h3', fontName='Carlito-Bold', fontSize=12.6, leading=16,
                         textColor=HexColor('#C2410C'), spaceBefore=8, spaceAfter=4)
S['bullet'] = ParagraphStyle('bullet', parent=S['body'], leftIndent=15, bulletIndent=3,
                             spaceAfter=3.2)
S['cell'] = ParagraphStyle('cell', fontName='Carlito', fontSize=9.4, leading=12.2, textColor=TEXT)
S['cellb'] = ParagraphStyle('cellb', parent=S['cell'], fontName='Carlito-Bold')
S['cellh'] = ParagraphStyle('cellh', parent=S['cell'], fontName='Carlito-Bold', textColor=white)
S['boxlabel'] = ParagraphStyle('boxlabel', fontName='Carlito-Bold', fontSize=9.6, leading=13)
S['boxbody'] = ParagraphStyle('boxbody', parent=S['body'], fontSize=10.3, leading=14.4, spaceAfter=0)
S['boxbullet'] = ParagraphStyle('boxbullet', parent=S['boxbody'], leftIndent=13, bulletIndent=2)
S['small'] = ParagraphStyle('small', parent=S['body'], fontSize=9, leading=12, textColor=MUTED)
S['toc0'] = ParagraphStyle('toc0', fontName='Carlito-Bold', fontSize=11.5, leading=15,
                           textColor=PRIMARY, spaceBefore=3, leftIndent=0)
S['toc1'] = ParagraphStyle('toc1', fontName='Carlito', fontSize=9.3, leading=10.7,
                           textColor=TEXT, leftIndent=16, spaceBefore=0.5)
S['parttitle'] = ParagraphStyle('parttitle', fontName='Carlito-Bold', fontSize=25, leading=29,
                                textColor=white)
S['partlabel'] = ParagraphStyle('partlabel', fontName='Carlito-Bold', fontSize=10.5, leading=13,
                                textColor=ACCENT)
S['partsub'] = ParagraphStyle('partsub', fontName='Carlito', fontSize=11.5, leading=15.5,
                              textColor=SOFT)


def strip_tags(t):
    return re.sub(r'<[^>]+>', '', t).replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')


def sym(ch, color=None):
    c = f' color="{color}"' if color else ''
    return f'<font name="Sym"{c}>{ch}</font>'


# ---------- Будівельні блоки ----------
class PartHeader(Table):
    def __init__(self, num, title, subtitle):
        word = 'ДОДАТОК' if num.isalpha() else 'ЧАСТИНА'
        rows = [[Paragraph(f'{word} {num}', S['partlabel'])],
                [Paragraph(title, S['parttitle'])]]
        if subtitle:
            rows.append([Paragraph(subtitle, S['partsub'])])
        super().__init__(rows, colWidths=[CW])
        self.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), PRIMARY),
            ('LEFTPADDING', (0, 0), (-1, -1), 22), ('RIGHTPADDING', (0, 0), (-1, -1), 22),
            ('TOPPADDING', (0, 0), (-1, 0), 20), ('BOTTOMPADDING', (0, -1), (-1, -1), 20),
            ('TOPPADDING', (0, 1), (-1, -1), 5),
            ('LINEBELOW', (0, -1), (-1, -1), 4, ACCENT),
            ('ROUNDEDCORNERS', [8, 8, 8, 8]),
        ]))
        self.toc_level = 0
        w2 = word.capitalize()
        self.toc_text = f'{w2} {num}. {strip_tags(title)}'
        self.running = f'{w2} {num} · {strip_tags(title)}'


def para(text, style='body'):
    return Paragraph(text, S[style])


def bullets(items, style='bullet', ordered=False, color='#F97316'):
    out = []
    for i, it in enumerate(items, 1):
        if ordered:
            out.append(Paragraph(it, S[style], bulletText=f'{i}.'))
        else:
            out.append(Paragraph(it, S[style], bulletText='•'))
    return out


def checklist(items):
    return [Paragraph(f'{sym("☐", "#312E81")}&nbsp;&nbsp;{it}', ParagraphStyle(
        'chk', parent=S['body'], leftIndent=16, firstLineIndent=-16, spaceAfter=3.5)) for it in items]


def box(kind, title, items):
    label, col, bg = BOX[kind]
    col_c, bg_c = HexColor(col), HexColor(bg)
    head = f'<font color="{col}">{label}</font>'
    if title:
        head += f'<font color="#1F2937">  ·  {title}</font>'
    rows = [[Paragraph(head, S['boxlabel'])]]
    for it in items:
        if isinstance(it, tuple) and it[0] in ('ul', 'ol'):
            for j, b in enumerate(it[1], 1):
                bt = f'{j}.' if it[0] == 'ol' else '•'
                rows.append([Paragraph(b, S['boxbullet'], bulletText=bt)])
        elif isinstance(it, tuple) and it[0] == 'chk':
            for b in it[1]:
                rows.append([Paragraph(f'{sym("☐", col)}&nbsp;&nbsp;{b}', S['boxbody'])])
        else:
            rows.append([Paragraph(it, S['boxbody'])])
    t = Table(rows, colWidths=[CW])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), bg_c),
        ('LINEBEFORE', (0, 0), (0, -1), 3.2, col_c),
        ('LEFTPADDING', (0, 0), (-1, -1), 13), ('RIGHTPADDING', (0, 0), (-1, -1), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 2.6), ('BOTTOMPADDING', (0, 0), (-1, -1), 2.6),
        ('TOPPADDING', (0, 0), (-1, 0), 9), ('BOTTOMPADDING', (0, -1), (-1, -1), 10),
    ]))
    return [Spacer(1, 4), KeepTogether([t]), Spacer(1, 9)]


def table(header, rows, widths, first_bold=True, font_size=None):
    tw = sum(widths)
    cw = [CW * w / tw for w in widths]
    cs, cb = S['cell'], S['cellb']
    if font_size:
        cs = ParagraphStyle('cs', parent=cs, fontSize=font_size, leading=font_size * 1.3)
        cb = ParagraphStyle('cb', parent=cs, fontName='Carlito-Bold')
    data = [[Paragraph(h, S['cellh']) for h in header]]
    groups = []
    for r in rows:
        if len(r) > 2 and r[0] == '' and all(c == '' for c in r[2:]):
            groups.append(len(data))
            data.append([Paragraph(f'<font color="#C2410C">{r[1]}</font>', cb)] + [''] * (len(r) - 1))
            continue
        data.append([Paragraph(str(c), cb if (i == 0 and first_bold) else cs) for i, c in enumerate(r)])
    t = Table(data, colWidths=cw, repeatRows=1)
    st = [
        ('BACKGROUND', (0, 0), (-1, 0), PRIMARY),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LINEBELOW', (0, 0), (-1, -1), 0.5, LINE),
        ('LEFTPADDING', (0, 0), (-1, -1), 6), ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 4.5), ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]
    for i in range(1, len(data)):
        if i in groups:
            st += [('SPAN', (0, i), (-1, i)), ('BACKGROUND', (0, i), (-1, i), HexColor('#FFF7ED')),
                   ('TOPPADDING', (0, i), (-1, i), 7)]
        elif i % 2 == 0:
            st.append(('BACKGROUND', (0, i), (-1, i), ZEBRA))
    t.setStyle(TableStyle(st))
    return [Spacer(1, 3), t, Spacer(1, 10)]


def qform_table(header, labels, widths, row_h=24):
    """Шаблон: заповнена ліва колонка, порожня права."""
    tw = sum(widths)
    cw = [CW * w / tw for w in widths]
    data = [[Paragraph(h, S['cellh']) for h in header]] + [[Paragraph(l, S['cellb']), ''] for l in labels]
    t = Table(data, colWidths=cw, rowHeights=[None] + [row_h] * len(labels), repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), PRIMARY),
        ('BACKGROUND', (0, 1), (0, -1), ZEBRA),
        ('GRID', (0, 0), (-1, -1), 0.5, HexColor('#CBD5E1')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    return [Spacer(1, 3), t, Spacer(1, 10)]


def form_table(header, nrows, widths, row_h=22):
    """Порожня таблиця-шаблон для заповнення."""
    tw = sum(widths)
    cw = [CW * w / tw for w in widths]
    data = [[Paragraph(h, S['cellh']) for h in header]] + [[''] * len(header) for _ in range(nrows)]
    t = Table(data, colWidths=cw, rowHeights=[None] + [row_h] * nrows, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), PRIMARY),
        ('GRID', (0, 0), (-1, -1), 0.5, HexColor('#CBD5E1')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    return [Spacer(1, 3), t, Spacer(1, 10)]


def quote(text, author):
    return box('quote', None, [f'<i>«{text}»</i>', f'<font color="#6B7280">— {author}</font>'])


# ---------- Діаграми ----------
def _txt(d, x, y, s, size=9, font='Carlito', color=TEXT, anchor='middle'):
    d.add(String(x, y, s, fontName=font, fontSize=size, fillColor=color, textAnchor=anchor))


def diag_attention():
    h = 108
    d = Drawing(CW, h)
    steps = [('1. ЗУПИНКА', 'Що це?', '0–1 с'), ('2. ЗАЧЕПЛЕННЯ', 'Це про мене?', '1–3 с'),
             ('3. УТРИМАННЯ', 'І що далі?', '3–15 с'), ('4. ДОВІРА', 'Чи можна вірити?', 'весь час'),
             ('5. ДІЯ', 'Що робити зараз?', 'фінал')]
    cols = ['#312E81', '#4338CA', '#6D28D9', '#C2410C', '#F97316']
    n = len(steps)
    w = CW / n
    tip = 12
    for i, (t, q, tm) in enumerate(steps):
        x0 = i * w
        pts = [x0, 30, x0 + w - 2, 30, x0 + w - 2 + tip, 62, x0 + w - 2, 94, x0, 94]
        if i > 0:
            pts += [x0 + tip, 62]
        d.add(Polygon(pts, fillColor=HexColor(cols[i]), strokeColor=None))
        cx = x0 + w / 2 + 4
        _txt(d, cx, 70, t, 9.2, 'Carlito-Bold', white)
        _txt(d, cx, 54, q, 9, 'Carlito', white)
        _txt(d, cx, 14, tm, 9, 'Carlito-Bold', MUTED)
    return [Spacer(1, 4), d, Spacer(1, 6)]


def diag_awareness():
    h = 96
    d = Drawing(CW, h)
    labels = [('Не знає', 'про проблему'), ('Знає', 'проблему'), ('Знає', 'рішення'),
              ('Знає', 'продукт'), ('Готовий', 'купити')]
    cols = ['#1E3A8A', '#3730A3', '#7E22CE', '#C2410C', '#EA580C']
    n = len(labels)
    w = CW / n
    for i, (a, b) in enumerate(labels):
        d.add(Rect(i * w + 1, 34, w - 2, 46, fillColor=HexColor(cols[i]), strokeColor=None))
        _txt(d, i * w + w / 2, 60, a, 10, 'Carlito-Bold', white)
        _txt(d, i * w + w / 2, 46, b, 9.2, 'Carlito', white)
    d.add(Line(4, 20, CW - 4, 20, strokeColor=MUTED, strokeWidth=1))
    d.add(Polygon([CW - 4, 20, CW - 12, 24, CW - 12, 16], fillColor=MUTED, strokeColor=None))
    _txt(d, 4, 6, 'ХОЛОДНА аудиторія: починаємо з людини та її світу', 8.8, 'Carlito-Bold', MUTED, 'start')
    _txt(d, CW - 4, 6, 'ГАРЯЧА: починаємо з продукту та оферу', 8.8, 'Carlito-Bold', MUTED, 'end')
    _txt(d, CW / 2, 86, 'Рівні усвідомленості (Ю. Шварц)', 9.5, 'Carlito-Bold', PRIMARY)
    return [Spacer(1, 4), d, Spacer(1, 8)]


def diag_hierarchy():
    levels = [('ОФЕР', 'що і на яких умовах', 1.00), ('КУТ', 'через який мотив заходимо', 0.88),
              ('КОНЦЕПЦІЯ', 'сюжет / формат подачі', 0.74), ('ХУК', 'перші 1–3 секунди', 0.60),
              ('ВИКОНАННЯ', 'актор, монтаж, кольори', 0.46), ('ВАРІАЦІЇ', 'дрібні зміни', 0.32)]
    rh = 22
    h = len(levels) * (rh + 4) + 24
    d = Drawing(CW, h)
    cols = ['#1E1B4B', '#312E81', '#4338CA', '#6D28D9', '#A855F7', '#C4B5FD']
    maxw = CW * 0.62
    for i, (a, b, k) in enumerate(levels):
        y = h - 22 - (i + 1) * (rh + 4) + 4
        w = maxw * k
        d.add(Rect(0, y, w, rh, fillColor=HexColor(cols[i]), strokeColor=None, rx=3, ry=3))
        _txt(d, 8, y + 7, a, 9.6, 'Carlito-Bold', white if i < 5 else PRIMARY, 'start')
        _txt(d, w + 8, y + 7, b, 9.4, 'Carlito', TEXT, 'start')
    _txt(d, 0, h - 12, 'Вплив на результат: від найбільшого (згори) до найменшого', 9.4,
         'Carlito-Bold', PRIMARY, 'start')
    return [Spacer(1, 4), d, Spacer(1, 8)]


def diag_stages():
    h = 92
    d = Drawing(CW, h)
    st = [('ЕТАП 1', 'Тест кутів', '3–5 кутів × 2–3 крео'), ('ЕТАП 2', 'Тест хуків', '5–10 хуків на кут-переможець'),
          ('ЕТАП 3', 'Формати', 'UGC, статика, довжина, актор'), ('ЕТАП 4', 'Масштаб + оновлення', 'щотижня нові ітерації')]
    cols = ['#312E81', '#6D28D9', '#C2410C', '#059669']
    n = len(st)
    gap = 14
    w = (CW - gap * (n - 1)) / n
    for i, (a, b, c) in enumerate(st):
        x = i * (w + gap)
        d.add(Rect(x, 14, w, 66, fillColor=HexColor(cols[i]), strokeColor=None, rx=5, ry=5))
        _txt(d, x + w / 2, 62, a, 8.8, 'Carlito-Bold', HexColor('#FED7AA'))
        _txt(d, x + w / 2, 45, b, 11, 'Carlito-Bold', white)
        _txt(d, x + w / 2, 27, c, 8.4, 'Carlito', white)
        if i < n - 1:
            ax = x + w + 2
            d.add(Polygon([ax, 52, ax + gap - 4, 47, ax, 42], fillColor=MUTED, strokeColor=None))
    return [Spacer(1, 4), d, Spacer(1, 8)]


def diag_cycle():
    h = 70
    d = Drawing(CW, h)
    st = ['Дослідження', 'Ідеї та хуки', 'Продакшн', 'Запуск тестів', 'Аналіз', 'Ітерації']
    n = len(st)
    gap = 10
    w = (CW - gap * (n - 1)) / n
    for i, s in enumerate(st):
        x = i * (w + gap)
        d.add(Rect(x, 22, w, 34, fillColor=HexColor('#EEF2FF'), strokeColor=PRIMARY2, strokeWidth=1, rx=4, ry=4))
        _txt(d, x + w / 2, 35, s, 9.4, 'Carlito-Bold', PRIMARY)
        if i < n - 1:
            ax = x + w + 1
            d.add(Polygon([ax, 43, ax + gap - 2, 39, ax, 35], fillColor=ACCENT, strokeColor=None))
    _txt(d, CW / 2, 6, 'повторюється щотижня: висновки з журналу тестів стають новими гіпотезами', 9, 'Carlito-Italic', MUTED)
    return [Spacer(1, 4), d, Spacer(1, 6)]


DIAGRAMS = {'attention': diag_attention, 'awareness': diag_awareness, 'hierarchy': diag_hierarchy,
            'stages': diag_stages, 'cycle': diag_cycle}


# ---------- Перетворення DSL у flowables ----------
def render(content):
    out = []
    for it in content:
        k = it[0]
        if k == 'h1':
            out += [PageBreak(), PartHeader(it[1], it[2], it[3] if len(it) > 3 else ''), Spacer(1, 14)]
        elif k == 'h1plain':
            out.append(PageBreak())
            p = Paragraph(it[1], S['h1plain'])
            p.toc_level = 0
            p.toc_text = strip_tags(it[1])
            p.running = strip_tags(it[1])
            out.append(p)
        elif k == 'h2':
            out.append(CondPageBreak(110))
            p = Paragraph(it[1], S['h2'])
            p.toc_level = 1
            p.toc_text = strip_tags(it[1])
            out.append(p)
        elif k == 'h3':
            out += [CondPageBreak(75), Paragraph(it[1], S['h3'])]
        elif k == 'p':
            out.append(Paragraph(it[1], S['body']))
        elif k == 'lead':
            out.append(Paragraph(it[1], S['lead']))
        elif k == 'small':
            out.append(Paragraph(it[1], S['small']))
        elif k == 'ul':
            out += bullets(it[1])
            out.append(Spacer(1, 4))
        elif k == 'ol':
            out += bullets(it[1], ordered=True)
            out.append(Spacer(1, 4))
        elif k == 'chk':
            out += checklist(it[1])
            out.append(Spacer(1, 4))
        elif k == 'box':
            out += box(it[1], it[2], it[3])
        elif k == 'table':
            fs = it[4] if len(it) > 4 else None
            out += table(it[1], it[2], it[3], font_size=fs)
        elif k == 'qform':
            out += qform_table(it[1], it[2], it[3], it[4] if len(it) > 4 else 24)
        elif k == 'form':
            out += form_table(it[1], it[2], it[3], it[4] if len(it) > 4 else 22)
        elif k == 'quote':
            out += quote(it[1], it[2])
        elif k == 'diagram':
            out += DIAGRAMS[it[1]]()
        elif k == 'pb':
            out.append(PageBreak())
        elif k == 'cpb':
            out.append(CondPageBreak(it[1]))
        elif k == 'sp':
            out.append(Spacer(1, it[1]))
        else:
            raise ValueError(k)
    return out


# ---------- Документ ----------
TITLE = 'КРЕАТИВИ, ЩО ВЛУЧАЮТЬ'


def draw_cover(c, doc):
    c.saveState()
    c.setFillColor(PRIMARY)
    c.rect(0, 0, PAGE_W, PAGE_H, stroke=0, fill=1)
    # декоративні кола
    c.setFillColor(ACCENT)
    c.circle(PAGE_W - 40, PAGE_H - 70, 150, stroke=0, fill=1)
    c.setFillColor(PRIMARY2)
    c.circle(PAGE_W - 40, PAGE_H - 70, 95, stroke=0, fill=1)
    c.setFillColor(HexColor('#4338CA'))
    c.circle(-50, 30, 100, stroke=0, fill=1)
    c.setStrokeColor(HexColor('#4F46E5'))
    c.setLineWidth(1)
    for i in range(7):
        c.line(LM, 330 - i * 9, LM + 160 - i * 18, 330 - i * 9)
    # тексти
    c.setFillColor(ACCENT)
    c.setFont('Carlito-Bold', 11)
    c.drawString(LM, PAGE_H - 250, 'НУТРА · СХУДНЕННЯ · ПРАКТИЧНИЙ ПОСІБНИК 2026')
    st = ParagraphStyle('ct', fontName='Carlito-Bold', fontSize=46, leading=48, textColor=white)
    p = Paragraph('Креативи,<br/>що влучають', st)
    w, h = p.wrap(PAGE_W - 2 * LM, 300)
    p.drawOn(c, LM, PAGE_H - 270 - h)
    y = PAGE_H - 270 - h - 20
    c.setFillColor(ACCENT)
    c.rect(LM, y, 70, 4, stroke=0, fill=1)
    st2 = ParagraphStyle('cs', fontName='Carlito', fontSize=15.5, leading=21, textColor=SOFT)
    p2 = Paragraph('Психологія аудиторії · Заголовки й тексти · Адаптація під будь-яке гео · '
                   'Система тестування · «Заборонені» тригери', st2)
    w2, h2 = p2.wrap(PAGE_W - 2 * LM - 60, 200)
    p2.drawOn(c, LM, y - 16 - h2)
    st3 = ParagraphStyle('cs3', fontName='Carlito', fontSize=11, leading=15, textColor=HexColor('#A5B4FC'))
    p3 = Paragraph('Для медіабаєрів, маркетологів, дизайнерів і креаторів у ніші схуднення, які хочуть '
                   'системно робити креативи, що зупиняють скрол, продають і не отримують бан.', st3)
    w3, h3 = p3.wrap(330, 200)
    p3.drawOn(c, LM, 120)
    c.setFillColor(HexColor('#818CF8'))
    c.setFont('Carlito', 9)
    c.drawString(LM, 60, '10 частин · 25 кутів · 30 хуків · 60 заголовків · 18 «заборонених» тригерів · план на 30 днів')
    c.restoreState()


def draw_page(c, doc):
    c.saveState()
    c.setStrokeColor(LINE)
    c.setLineWidth(0.6)
    c.line(LM, PAGE_H - 40, PAGE_W - RM, PAGE_H - 40)
    c.setFont('Carlito-Bold', 8)
    c.setFillColor(PRIMARY2)
    c.drawString(LM, PAGE_H - 34, TITLE)
    c.setFont('Carlito', 8)
    c.setFillColor(MUTED)
    run = doc.current_part
    if len(run) > 75:
        run = run[:72] + '…'
    c.drawRightString(PAGE_W - RM, PAGE_H - 34, run)
    c.line(LM, 38, PAGE_W - RM, 38)
    c.setFont('Carlito-Bold', 9)
    c.setFillColor(PRIMARY)
    c.drawRightString(PAGE_W - RM, 24, str(doc.page))
    c.setFillColor(ACCENT)
    c.rect(PAGE_W - RM - 26, 22, 3, 11, stroke=0, fill=1)
    c.restoreState()


class GuideDoc(BaseDocTemplate):
    def __init__(self, fn, **kw):
        super().__init__(fn, pagesize=A4, leftMargin=LM, rightMargin=RM, topMargin=TM,
                         bottomMargin=BM, **kw)
        cover = Frame(0, 0, PAGE_W, PAGE_H, id='cover', leftPadding=0, rightPadding=0,
                      topPadding=0, bottomPadding=0)
        body = Frame(LM, BM, CW, PAGE_H - TM - BM, id='body', leftPadding=0, rightPadding=0,
                     topPadding=0, bottomPadding=0)
        self.addPageTemplates([PageTemplate('cover', [cover], onPage=draw_cover),
                               PageTemplate('body', [body], onPageEnd=draw_page)])
        self.current_part = ''
        self._k = 0

    def beforeDocument(self):
        self.current_part = ''
        self._k = 0

    def afterFlowable(self, f):
        lvl = getattr(f, 'toc_level', None)
        if lvl is None:
            return
        if lvl == 0:
            self.current_part = getattr(f, 'running', f.toc_text)
        self._k += 1
        key = f'k{self._k}'
        self.canv.bookmarkPage(key)
        self.canv.addOutlineEntry(f.toc_text, key, level=lvl, closed=True)
        self.notify('TOCEntry', (lvl, f.toc_text, self.page, key))


def make_toc():
    toc = TableOfContents()
    toc.levelStyles = [S['toc0'], S['toc1']]
    toc.dotsMinLevel = 0
    return toc
