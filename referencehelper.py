#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2024/3/15 14:20
# @Author : Jiyun Zhang

import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def txt_to_word(input_txt, output_doc):
    with open(input_txt, 'r') as file:
        lines = file.read().split('\n\n')

    doc = Document()

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)

    lines.sort(key=lambda x: (x.split('\n')[1], x.split('\n')[2]))

    for line in lines:
        parts = line.split('\n')
        author, initial, year, title, url = parts[1], parts[2], parts[3], parts[4], parts[5]

        italic = False
        if title.startswith('#x'):
            italic = True
            title = title[2:].strip()

        para = doc.add_paragraph()
        para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        para.add_run(f'{author}, {initial} {year}, ')
        para.add_run(f'{title} ').italic = italic
        para.add_run(f'viewed {datetime.datetime.now().strftime("%d %B %Y")}, ')
        para.add_run(f'<{url}>.')

    doc.save(output_doc)

txt_to_word('input.txt', 'output.docx')