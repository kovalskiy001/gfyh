# -*- coding: utf-8 -*-
"""Збирає PDF-посібник. Запуск: python3 build.py [шлях_до_pdf]"""
import os
import sys

from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import NextPageTemplate, PageBreak, Paragraph

from engine import GuideDoc, make_toc, render, S
from content_a import INTRO, PART1, PART2
from content_b import EMOTIONS, PART3, PART4, PART5
from content_c import PART6, PART7
from content_d import PART8, PART9, PART10, APPX

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '..', 'Kreatyvy-Nutra-Shudnennya.pdf')


def main():
    doc = GuideDoc(OUT, title='Креативи, що влучають',
                   subject='Нутра на схуднення: психологія аудиторії, заголовки, адаптація під гео, тестування креативів')
    toc_title = ParagraphStyle('toct', parent=S['h1plain'], fontSize=20, leading=22, spaceAfter=2)
    story = [NextPageTemplate('body'), PageBreak(), Paragraph('Зміст', toc_title), make_toc()]
    content = INTRO + PART1 + EMOTIONS + PART2 + PART3 + PART4 + PART5 + PART6 + PART7 + PART8 + PART9 + PART10 + APPX
    story += render(content)
    doc.multiBuild(story)
    print('OK:', os.path.abspath(OUT))


if __name__ == '__main__':
    main()
