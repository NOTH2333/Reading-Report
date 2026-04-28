#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import fitz
import json
import sys

pdf_path = '/g/OBSIDIAN/天才的回声：经济学大师与他们塑造的世界 ([美]拖德·布赫霍尔茨) (z-library.sk, 1lib.sk, z-lib.sk).pdf'

doc = fitz.open(pdf_path)
print(f'总页数: {doc.page_count}')

# 提取目录
toc = doc.get_toc()
print('\n=== 目录 ===')
if toc:
    for item in toc:
        level, title, page = item
        print(f'L{level}|{title}|p{page}')
else:
    print('无内置目录，扫描前30页寻找目录页')
    for i in range(min(30, doc.page_count)):
        page = doc[i]
        text = page.get_text()
        if text.strip():
            print(f'\n--- 第{i+1}页 ---')
            print(text[:800])
            print('...')
