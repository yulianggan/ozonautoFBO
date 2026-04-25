# encoding='utf-8

# @Time: 2024-05-18
# @File: %
#!/usr/bin/env
from icecream import ic
import os

import fitz  # PyMuPDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import os
from PyPDF2 import PdfMerger

def merge_pdfs(input_folder, output_pdf_path):
    # 创建一个PdfMerger对象
    merger = PdfMerger()
    
    # 遍历文件夹中的所有PDF文件
    for item in sorted(os.listdir(input_folder)):
        if item.endswith('.pdf'):
            file_path = os.path.join(input_folder, item)
            # 将PDF文件添加到合并对象中
            merger.append(file_path)
    
    # 将所有PDF文件合并成一个文件
    merger.write(output_pdf_path)
    merger.close()
    return output_pdf_path

# input_folder = './test'  # 替换为你的PDF文件夹路径
# output_pdf_path = './merged_output.pdf'  # 合并后的PDF文件路径
#
# merge_pdfs(input_folder, output_pdf_path)
#
# exit()
import polars as pl

def add_text_to_pdf(input_pdf_path, output_pdf_path, input_infos):
    # 打开原始PDF文件
    document = fitz.open(input_pdf_path)
    df = pl.read_excel(input_infos,sheet_name='91775197-pallet')
    # ic(df)
    df = df.select(['货位 ID', '箱唛'])
    for i in range(len(df)):
        search_text = str(df['货位 ID'][i])
        ic(search_text)
        text_to_add = str(df['箱唛'][i]).replace("&","\n")
        ic(text_to_add)

        for page_num in range(len(document)):
            page = document.load_page(page_num)
            text_instances = page.search_for(search_text)

            for inst in text_instances:
                # 获取搜索文本的矩形区域
                x0, y0, x1, y1 = inst

                # 在搜索文本下面的空白位置添加新文本
                page.insert_text((x0 - 80, y1 - 80), text_to_add, fontsize=10, color=(0, 0, 0))

    # 保存修改后的PDF文件
    document.save(output_pdf_path)


# input_pdf_path = './merged_output.pdf'
input_pdf_path = merge_pdfs('./merged', './merged_output.pdf') # merge pdfs
output_pdf_path = 'output.pdf'
input_infos = './汇总的装箱单分配仓库-模板.xlsm'
# search_text = '24586517185000'  # 替换为你想要搜索的文本
# text_to_add = '1-A030'

add_text_to_pdf(input_pdf_path, output_pdf_path, input_infos)


