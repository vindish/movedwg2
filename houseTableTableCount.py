#!/usr/bin/python
# -*- coding: utf-8 -*-
import os, xlwt, time
from docx import Document  # 导入库
nowtime = time.strftime("%H%M%S")
#rootdir = ((os.getcwd()).replace('\\', '/').decode(encoding="gb2312", errors="strict"))
rootdir = ("d:\Data\docx")
lists = os.listdir(rootdir)  # 列出文件夹下所有的目录与文件
docxlist = []
for ListName in lists:
    fname, ext = os.path.splitext(ListName)
    if ext == '.docx':
        docxlist.append(ListName)
ListName = docxlist
f = xlwt.Workbook()  # 创建工作簿
sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet
t = []
cell_lenths_list = []
for i in range(0, len(ListName)):
    path = os.path.join(rootdir, ListName[i])
    if os.path.isfile(path):  # 如果是文件夹就继续打开
        document = Document(path)  # 读入文件
        tables = document.tables  # 获取文件中的表格集
#       table = tables[0]  # 获取文件中的第一个表格
        celllist = []
        index, key, value = [], [], []
        table_index = 0
        for table in tables:
            table_index += 1
            row_index = 0
            for row in table.rows:
                row_index += 1
                for cell in row.cells:
                    celllist.append(cell.text)
            print (i+1, ListName[i], len(celllist), "开始写入")
            len_celli = len(celllist)
            data = []
            if len_celli == 513:
                R0C0 = ("513"+table.cell(0, 0).text)
                data.append(R0C0)
                R1C0 = table.cell(1, 0).text
                data.append(R1C0)
                R2C1 = table.cell(2, 1).text
                data.append(R2C1)
                R3C1 = table.cell(3, 1).text
                data.append(R3C1)
                R3C13 = table.cell(3, 13).text
                data.append(R3C13)
                R4C1 = table.cell(4, 1).text
                data.append(R4C1)
                R4C9 = table.cell(4, 9).text
                data.append(R4C9)
                R5C9 = table.cell(5, 9).text
                data.append(R5C9)
                R6C1 = table.cell(6, 1).text
                data.append(R6C1)
                R6C7 = table.cell(6, 7).text
                data.append(R6C7)
                R6C10 = table.cell(6, 10).text
                data.append(R6C10)
                R7C1 = table.cell(7, 1).text
                data.append(R7C1)
                R7C7 = table.cell(7, 7).text
                data.append(R7C7)
                R8C1 = table.cell(8, 1).text
                data.append(R8C1)
                R8C7 = table.cell(8, 7).text
                data.append(R8C7)
                R9C1 = table.cell(9, 1).text
                data.append(R9C1)
                R9C7 = table.cell(9, 7).text
                data.append(R9C7)
                R12C1 = table.cell(12, 1).text
                data.append(R12C1)
                R12C2 = table.cell(12, 2).text
                data.append(R12C2)
                R12C3 = table.cell(12, 3).text
                data.append(R12C3)
                R12C4 = table.cell(12, 4).text
                data.append(R12C4)
                R12C5 = table.cell(12, 5).text
                data.append(R12C5)
                R12C6 = table.cell(12, 6).text
                data.append(R12C6)
                R12C7 = table.cell(12, 7).text
                data.append(R12C7)
                R12C8 = table.cell(12, 8).text
                data.append(R12C8)
                R12C9 = table.cell(12, 9).text
                data.append(R12C9)
                R12C10 = table.cell(12, 10).text
                data.append(R12C10)
                R12C12 = table.cell(12, 12).text
                data.append(R12C12)
                R12C14 = table.cell(12, 14).text
                data.append(R12C14)
                R12C15 = table.cell(12, 15).text
                data.append(R12C15)
                R12C16 = table.cell(12, 16).text
                data.append(R12C16)
                R12C17 = table.cell(12, 17).text
                data.append(R12C17)
                R12C18 = table.cell(12, 18).text
                data.append(R12C18)
                R13C1 = table.cell(13, 1).text
                data.append(R13C1)
                R13C2 = table.cell(13, 2).text
                data.append(R13C2)
                R13C3 = table.cell(13, 3).text
                data.append(R13C3)
                R13C4 = table.cell(13, 4).text
                data.append(R13C4)
                R13C5 = table.cell(13, 5).text
                data.append(R13C5)
                R13C6 = table.cell(13, 6).text
                data.append(R13C6)
                R13C7 = table.cell(13, 7).text
                data.append(R13C7)
                R13C8 = table.cell(13, 8).text
                data.append(R13C8)
                R13C9 = table.cell(13, 9).text
                data.append(R13C9)
                R13C10 = table.cell(13, 10).text
                data.append(R13C10)
                R13C12 = table.cell(13, 12).text
                data.append(R13C12)
                R13C14 = table.cell(13, 14).text
                data.append(R13C14)
                R13C15 = table.cell(13, 15).text
                data.append(R13C15)
                R13C16 = table.cell(13, 16).text
                data.append(R13C16)
                R13C17 = table.cell(13, 17).text
                data.append(R13C17)
                R13C18 = table.cell(13, 18).text
                data.append(R13C18)
                R14C1 = table.cell(14, 1).text
                data.append(R14C1)
                R14C2 = table.cell(14, 2).text
                data.append(R14C2)
                R14C3 = table.cell(14, 3).text
                data.append(R14C3)
                R14C4 = table.cell(14, 4).text
                data.append(R14C4)
                R14C5 = table.cell(14, 5).text
                data.append(R14C5)
                R14C6 = table.cell(14, 6).text
                data.append(R14C6)
                R14C7 = table.cell(14, 7).text
                data.append(R14C7)
                R14C8 = table.cell(14, 8).text
                data.append(R14C8)
                R14C9 = table.cell(14, 9).text
                data.append(R14C9)
                R14C10 = table.cell(14, 10).text
                data.append(R14C10)
                R14C12 = table.cell(14, 12).text
                data.append(R14C12)
                R14C14 = table.cell(14, 14).text
                data.append(R14C14)
                R14C15 = table.cell(14, 15).text
                data.append(R14C15)
                R14C16 = table.cell(14, 16).text
                data.append(R14C16)
                R14C17 = table.cell(14, 17).text
                data.append(R14C17)
                R14C18 = table.cell(14, 18).text
                data.append(R14C18)
                R15C1 = table.cell(15, 1).text
                data.append(R15C1)
                R15C2 = table.cell(15, 2).text
                data.append(R15C2)
                R15C3 = table.cell(15, 3).text
                data.append(R15C3)
                R15C4 = table.cell(15, 4).text
                data.append(R15C4)
                R15C5 = table.cell(15, 5).text
                data.append(R15C5)
                R15C6 = table.cell(15, 6).text
                data.append(R15C6)
                R15C7 = table.cell(15, 7).text
                data.append(R15C7)
                R15C8 = table.cell(15, 8).text
                data.append(R15C8)
                R15C9 = table.cell(15, 9).text
                data.append(R15C9)
                R15C10 = table.cell(15, 10).text
                data.append(R15C10)
                R15C12 = table.cell(15, 12).text
                data.append(R15C12)
                R15C14 = table.cell(15, 14).text
                data.append(R15C14)
                R15C15 = table.cell(15, 15).text
                data.append(R15C15)
                R15C16 = table.cell(15, 16).text
                data.append(R15C16)
                R15C17 = table.cell(15, 17).text
                data.append(R15C17)
                R15C18 = table.cell(15, 18).text
                data.append(R15C18)
                R16C1 = table.cell(16, 1).text
                data.append(R16C1)
                R16C2 = table.cell(16, 2).text
                data.append(R16C2)
                R16C3 = table.cell(16, 3).text
                data.append(R16C3)
                R16C4 = table.cell(16, 4).text
                data.append(R16C4)
                R16C5 = table.cell(16, 5).text
                data.append(R16C5)
                R16C6 = table.cell(16, 6).text
                data.append(R16C6)
                R16C7 = table.cell(16, 7).text
                data.append(R16C7)
                R16C8 = table.cell(16, 8).text
                data.append(R16C8)
                R16C9 = table.cell(16, 9).text
                data.append(R16C9)
                R16C10 = table.cell(16, 10).text
                data.append(R16C10)
                R16C12 = table.cell(16, 12).text
                data.append(R16C12)
                R16C14 = table.cell(16, 14).text
                data.append(R16C14)
                R16C15 = table.cell(16, 15).text
                data.append(R16C15)
                R16C16 = table.cell(16, 16).text
                data.append(R16C16)
                R16C17 = table.cell(16, 17).text
                data.append(R16C17)
                R16C18 = table.cell(16, 18).text
                data.append(R16C18)
                R17C1 = table.cell(17, 1).text
                data.append(R17C1)
                R17C2 = table.cell(17, 2).text
                data.append(R17C2)
                R17C3 = table.cell(17, 3).text
                data.append(R17C3)
                R17C4 = table.cell(17, 4).text
                data.append(R17C4)
                R17C5 = table.cell(17, 5).text
                data.append(R17C5)
                R17C6 = table.cell(17, 6).text
                data.append(R17C6)
                R17C7 = table.cell(17, 7).text
                data.append(R17C7)
                R17C8 = table.cell(17, 8).text
                data.append(R17C8)
                R17C9 = table.cell(17, 9).text
                data.append(R17C9)
                R17C10 = table.cell(17, 10).text
                data.append(R17C10)
                R17C12 = table.cell(17, 12).text
                data.append(R17C12)
                R17C14 = table.cell(17, 14).text
                data.append(R17C14)
                R17C15 = table.cell(17, 15).text
                data.append(R17C15)
                R17C16 = table.cell(17, 16).text
                data.append(R17C16)
                R17C17 = table.cell(17, 17).text
                data.append(R17C17)
                R17C18 = table.cell(17, 18).text
                data.append(R17C18)
                R18C1 = table.cell(18, 1).text
                data.append(R18C1)
                R18C2 = table.cell(18, 2).text
                data.append(R18C2)
                R18C3 = table.cell(18, 3).text
                data.append(R18C3)
                R18C4 = table.cell(18, 4).text
                data.append(R18C4)
                R18C5 = table.cell(18, 5).text
                data.append(R18C5)
                R18C6 = table.cell(18, 6).text
                data.append(R18C6)
                R18C7 = table.cell(18, 7).text
                data.append(R18C7)
                R18C8 = table.cell(18, 8).text
                data.append(R18C8)
                R18C9 = table.cell(18, 9).text
                data.append(R18C9)
                R18C10 = table.cell(18, 10).text
                data.append(R18C10)
                R18C12 = table.cell(18, 12).text
                data.append(R18C12)
                R18C14 = table.cell(18, 14).text
                data.append(R18C14)
                R18C15 = table.cell(18, 15).text
                data.append(R18C15)
                R18C16 = table.cell(18, 16).text
                data.append(R18C16)
                R18C17 = table.cell(18, 17).text
                data.append(R18C17)
                R18C18 = table.cell(18, 18).text
                data.append(R18C18)
                R19C1 = table.cell(19, 1).text
                data.append(R19C1)
                R19C2 = table.cell(19, 2).text
                data.append(R19C2)
                R19C3 = table.cell(19, 3).text
                data.append(R19C3)
                R19C4 = table.cell(19, 4).text
                data.append(R19C4)
                R19C5 = table.cell(19, 5).text
                data.append(R19C5)
                R19C6 = table.cell(19, 6).text
                data.append(R19C6)
                R19C7 = table.cell(19, 7).text
                data.append(R19C7)
                R19C8 = table.cell(19, 8).text
                data.append(R19C8)
                R19C9 = table.cell(19, 9).text
                data.append(R19C9)
                R19C10 = table.cell(19, 10).text
                data.append(R19C10)
                R19C12 = table.cell(19, 12).text
                data.append(R19C12)
                R19C14 = table.cell(19, 14).text
                data.append(R19C14)
                R19C15 = table.cell(19, 15).text
                data.append(R19C15)
                R19C16 = table.cell(19, 16).text
                data.append(R19C16)
                R19C17 = table.cell(19, 17).text
                data.append(R19C17)
                R19C18 = table.cell(19, 18).text
                data.append(R19C18)
                R20C1 = table.cell(20, 1).text
                data.append(R20C1)
                R20C2 = table.cell(20, 2).text
                data.append(R20C2)
                R20C3 = table.cell(20, 3).text
                data.append(R20C3)
                R20C4 = table.cell(20, 4).text
                data.append(R20C4)
                R20C5 = table.cell(20, 5).text
                data.append(R20C5)
                R20C6 = table.cell(20, 6).text
                data.append(R20C6)
                R20C7 = table.cell(20, 7).text
                data.append(R20C7)
                R20C8 = table.cell(20, 8).text
                data.append(R20C8)
                R20C9 = table.cell(20, 9).text
                data.append(R20C9)
                R20C10 = table.cell(20, 10).text
                data.append(R20C10)
                R20C12 = table.cell(20, 12).text
                data.append(R20C12)
                R20C14 = table.cell(20, 14).text
                data.append(R20C14)
                R20C15 = table.cell(20, 15).text
                data.append(R20C15)
                R20C16 = table.cell(20, 16).text
                data.append(R20C16)
                R20C17 = table.cell(20, 17).text
                data.append(R20C17)
                R20C18 = table.cell(20, 18).text
                data.append(R20C18)
                R21C1 = table.cell(21, 1).text
                data.append(R21C1)
                R21C2 = table.cell(21, 2).text
                data.append(R21C2)
                R21C3 = table.cell(21, 3).text
                data.append(R21C3)
                R21C4 = table.cell(21, 4).text
                data.append(R21C4)
                R21C5 = table.cell(21, 5).text
                data.append(R21C5)
                R21C6 = table.cell(21, 6).text
                data.append(R21C6)
                R21C7 = table.cell(21, 7).text
                data.append(R21C7)
                R21C8 = table.cell(21, 8).text
                data.append(R21C8)
                R21C9 = table.cell(21, 9).text
                data.append(R21C9)
                R21C10 = table.cell(21, 10).text
                data.append(R21C10)
                R21C12 = table.cell(21, 12).text
                data.append(R21C12)
                R21C14 = table.cell(21, 14).text
                data.append(R21C14)
                R21C15 = table.cell(21, 15).text
                data.append(R21C15)
                R21C16 = table.cell(21, 16).text
                data.append(R21C16)
                R21C17 = table.cell(21, 17).text
                data.append(R21C17)
                R21C18 = table.cell(21, 18).text
                data.append(R21C18)
                R22C1 = table.cell(22, 1).text
                data.append(R22C1)
                R22C2 = table.cell(22, 2).text
                data.append(R22C2)
                R22C3 = table.cell(22, 3).text
                data.append(R22C3)
                R22C4 = table.cell(22, 4).text
                data.append(R22C4)
                R22C5 = table.cell(22, 5).text
                data.append(R22C5)
                R22C6 = table.cell(22, 6).text
                data.append(R22C6)
                R22C7 = table.cell(22, 7).text
                data.append(R22C7)
                R22C8 = table.cell(22, 8).text
                data.append(R22C8)
                R22C9 = table.cell(22, 9).text
                data.append(R22C9)
                R22C10 = table.cell(22, 10).text
                data.append(R22C10)
                R22C12 = table.cell(22, 12).text
                data.append(R22C12)
                R22C14 = table.cell(22, 14).text
                data.append(R22C14)
                R22C15 = table.cell(22, 15).text
                data.append(R22C15)
                R22C16 = table.cell(22, 16).text
                data.append(R22C16)
                R22C17 = table.cell(22, 17).text
                data.append(R22C17)
                R22C18 = table.cell(22, 18).text
                data.append(R22C18)
                R23C1 = table.cell(23, 1).text
                data.append(R23C1)
                R23C10 = table.cell(23, 10).text
                data.append(R23C10)
                R24C10 = table.cell(24, 10).text
                data.append(R24C10)
                R25C0 = table.cell(25, 0).text
                data.append(R25C0)
                R26C0 = table.cell(26, 0).text
                data.append(R26C0)
            elif len_celli == 494:
                R0C0 = ("494"+table.cell(0, 0).text)
                data.append(R0C0)
                R1C0 = table.cell(1, 0).text
                data.append(R1C0)
                R2C1 = table.cell(2, 1).text
                data.append(R2C1)
                R3C1 = table.cell(3, 1).text
                data.append(R3C1)
                R3C13 = table.cell(3, 13).text
                data.append(R3C13)
                R4C1 = table.cell(4, 1).text
                data.append(R4C1)
                R4C9 = table.cell(4, 9).text
                data.append(R4C9)
                R5C9 = table.cell(5, 9).text
                data.append(R5C9)
                R6C1 = table.cell(6, 1).text
                data.append(R6C1)
                R6C7 = table.cell(6, 7).text
                data.append(R6C7)
                R6C10 = table.cell(6, 10).text
                data.append(R6C10)
                R7C1 = table.cell(7, 1).text
                data.append(R7C1)
                R7C7 = table.cell(7, 7).text
                data.append(R7C7)
                R8C1 = table.cell(8, 1).text
                data.append(R8C1)
                R8C7 = table.cell(8, 7).text
                data.append(R8C7)
                R9C1 = table.cell(9, 1).text
                data.append(R9C1)
                R9C7 = table.cell(9, 7).text
                data.append(R9C7)
                R12C1 = table.cell(12, 1).text
                data.append(R12C1)
                R12C2 = table.cell(12, 2).text
                data.append(R12C2)
                R12C3 = table.cell(12, 3).text
                data.append(R12C3)
                R12C4 = table.cell(12, 4).text
                data.append(R12C4)
                R12C5 = table.cell(12, 5).text
                data.append(R12C5)
                R12C6 = table.cell(12, 6).text
                data.append(R12C6)
                R12C7 = table.cell(12, 7).text
                data.append(R12C7)
                R12C8 = table.cell(12, 8).text
                data.append(R12C8)
                R12C9 = table.cell(12, 9).text
                data.append(R12C9)
                R12C10 = table.cell(12, 10).text
                data.append(R12C10)
                R12C12 = table.cell(12, 12).text
                data.append(R12C12)
                R12C14 = table.cell(12, 14).text
                data.append(R12C14)
                R12C15 = table.cell(12, 15).text
                data.append(R12C15)
                R12C16 = table.cell(12, 16).text
                data.append(R12C16)
                R12C17 = table.cell(12, 17).text
                data.append(R12C17)
                R12C18 = table.cell(12, 18).text
                data.append(R12C18)
                R13C1 = table.cell(13, 1).text
                data.append(R13C1)
                R13C2 = table.cell(13, 2).text
                data.append(R13C2)
                R13C3 = table.cell(13, 3).text
                data.append(R13C3)
                R13C4 = table.cell(13, 4).text
                data.append(R13C4)
                R13C5 = table.cell(13, 5).text
                data.append(R13C5)
                R13C6 = table.cell(13, 6).text
                data.append(R13C6)
                R13C7 = table.cell(13, 7).text
                data.append(R13C7)
                R13C8 = table.cell(13, 8).text
                data.append(R13C8)
                R13C9 = table.cell(13, 9).text
                data.append(R13C9)
                R13C10 = table.cell(13, 10).text
                data.append(R13C10)
                R13C12 = table.cell(13, 12).text
                data.append(R13C12)
                R13C14 = table.cell(13, 14).text
                data.append(R13C14)
                R13C15 = table.cell(13, 15).text
                data.append(R13C15)
                R13C16 = table.cell(13, 16).text
                data.append(R13C16)
                R13C17 = table.cell(13, 17).text
                data.append(R13C17)
                R13C18 = table.cell(13, 18).text
                data.append(R13C18)
                R14C1 = table.cell(14, 1).text
                data.append(R14C1)
                R14C2 = table.cell(14, 2).text
                data.append(R14C2)
                R14C3 = table.cell(14, 3).text
                data.append(R14C3)
                R14C4 = table.cell(14, 4).text
                data.append(R14C4)
                R14C5 = table.cell(14, 5).text
                data.append(R14C5)
                R14C6 = table.cell(14, 6).text
                data.append(R14C6)
                R14C7 = table.cell(14, 7).text
                data.append(R14C7)
                R14C8 = table.cell(14, 8).text
                data.append(R14C8)
                R14C9 = table.cell(14, 9).text
                data.append(R14C9)
                R14C10 = table.cell(14, 10).text
                data.append(R14C10)
                R14C12 = table.cell(14, 12).text
                data.append(R14C12)
                R14C14 = table.cell(14, 14).text
                data.append(R14C14)
                R14C15 = table.cell(14, 15).text
                data.append(R14C15)
                R14C16 = table.cell(14, 16).text
                data.append(R14C16)
                R14C17 = table.cell(14, 17).text
                data.append(R14C17)
                R14C18 = table.cell(14, 18).text
                data.append(R14C18)
                R15C1 = table.cell(15, 1).text
                data.append(R15C1)
                R15C2 = table.cell(15, 2).text
                data.append(R15C2)
                R15C3 = table.cell(15, 3).text
                data.append(R15C3)
                R15C4 = table.cell(15, 4).text
                data.append(R15C4)
                R15C5 = table.cell(15, 5).text
                data.append(R15C5)
                R15C6 = table.cell(15, 6).text
                data.append(R15C6)
                R15C7 = table.cell(15, 7).text
                data.append(R15C7)
                R15C8 = table.cell(15, 8).text
                data.append(R15C8)
                R15C9 = table.cell(15, 9).text
                data.append(R15C9)
                R15C10 = table.cell(15, 10).text
                data.append(R15C10)
                R15C12 = table.cell(15, 12).text
                data.append(R15C12)
                R15C14 = table.cell(15, 14).text
                data.append(R15C14)
                R15C15 = table.cell(15, 15).text
                data.append(R15C15)
                R15C16 = table.cell(15, 16).text
                data.append(R15C16)
                R15C17 = table.cell(15, 17).text
                data.append(R15C17)
                R15C18 = table.cell(15, 18).text
                data.append(R15C18)
                R16C1 = table.cell(16, 1).text
                data.append(R16C1)
                R16C2 = table.cell(16, 2).text
                data.append(R16C2)
                R16C3 = table.cell(16, 3).text
                data.append(R16C3)
                R16C4 = table.cell(16, 4).text
                data.append(R16C4)
                R16C5 = table.cell(16, 5).text
                data.append(R16C5)
                R16C6 = table.cell(16, 6).text
                data.append(R16C6)
                R16C7 = table.cell(16, 7).text
                data.append(R16C7)
                R16C8 = table.cell(16, 8).text
                data.append(R16C8)
                R16C9 = table.cell(16, 9).text
                data.append(R16C9)
                R16C10 = table.cell(16, 10).text
                data.append(R16C10)
                R16C12 = table.cell(16, 12).text
                data.append(R16C12)
                R16C14 = table.cell(16, 14).text
                data.append(R16C14)
                R16C15 = table.cell(16, 15).text
                data.append(R16C15)
                R16C16 = table.cell(16, 16).text
                data.append(R16C16)
                R16C17 = table.cell(16, 17).text
                data.append(R16C17)
                R16C18 = table.cell(16, 18).text
                data.append(R16C18)
                R17C1 = table.cell(17, 1).text
                data.append(R17C1)
                R17C2 = table.cell(17, 2).text
                data.append(R17C2)
                R17C3 = table.cell(17, 3).text
                data.append(R17C3)
                R17C4 = table.cell(17, 4).text
                data.append(R17C4)
                R17C5 = table.cell(17, 5).text
                data.append(R17C5)
                R17C6 = table.cell(17, 6).text
                data.append(R17C6)
                R17C7 = table.cell(17, 7).text
                data.append(R17C7)
                R17C8 = table.cell(17, 8).text
                data.append(R17C8)
                R17C9 = table.cell(17, 9).text
                data.append(R17C9)
                R17C10 = table.cell(17, 10).text
                data.append(R17C10)
                R17C12 = table.cell(17, 12).text
                data.append(R17C12)
                R17C14 = table.cell(17, 14).text
                data.append(R17C14)
                R17C15 = table.cell(17, 15).text
                data.append(R17C15)
                R17C16 = table.cell(17, 16).text
                data.append(R17C16)
                R17C17 = table.cell(17, 17).text
                data.append(R17C17)
                R17C18 = table.cell(17, 18).text
                data.append(R17C18)
                R18C1 = table.cell(18, 1).text
                data.append(R18C1)
                R18C2 = table.cell(18, 2).text
                data.append(R18C2)
                R18C3 = table.cell(18, 3).text
                data.append(R18C3)
                R18C4 = table.cell(18, 4).text
                data.append(R18C4)
                R18C5 = table.cell(18, 5).text
                data.append(R18C5)
                R18C6 = table.cell(18, 6).text
                data.append(R18C6)
                R18C7 = table.cell(18, 7).text
                data.append(R18C7)
                R18C8 = table.cell(18, 8).text
                data.append(R18C8)
                R18C9 = table.cell(18, 9).text
                data.append(R18C9)
                R18C10 = table.cell(18, 10).text
                data.append(R18C10)
                R18C12 = table.cell(18, 12).text
                data.append(R18C12)
                R18C14 = table.cell(18, 14).text
                data.append(R18C14)
                R18C15 = table.cell(18, 15).text
                data.append(R18C15)
                R18C16 = table.cell(18, 16).text
                data.append(R18C16)
                R18C17 = table.cell(18, 17).text
                data.append(R18C17)
                R18C18 = table.cell(18, 18).text
                data.append(R18C18)
                R19C1 = table.cell(19, 1).text
                data.append(R19C1)
                R19C2 = table.cell(19, 2).text
                data.append(R19C2)
                R19C3 = table.cell(19, 3).text
                data.append(R19C3)
                R19C4 = table.cell(19, 4).text
                data.append(R19C4)
                R19C5 = table.cell(19, 5).text
                data.append(R19C5)
                R19C6 = table.cell(19, 6).text
                data.append(R19C6)
                R19C7 = table.cell(19, 7).text
                data.append(R19C7)
                R19C8 = table.cell(19, 8).text
                data.append(R19C8)
                R19C9 = table.cell(19, 9).text
                data.append(R19C9)
                R19C10 = table.cell(19, 10).text
                data.append(R19C10)
                R19C12 = table.cell(19, 12).text
                data.append(R19C12)
                R19C14 = table.cell(19, 14).text
                data.append(R19C14)
                R19C15 = table.cell(19, 15).text
                data.append(R19C15)
                R19C16 = table.cell(19, 16).text
                data.append(R19C16)
                R19C17 = table.cell(19, 17).text
                data.append(R19C17)
                R19C18 = table.cell(19, 18).text
                data.append(R19C18)
                R20C1 = table.cell(20, 1).text
                data.append(R20C1)
                R20C2 = table.cell(20, 2).text
                data.append(R20C2)
                R20C3 = table.cell(20, 3).text
                data.append(R20C3)
                R20C4 = table.cell(20, 4).text
                data.append(R20C4)
                R20C5 = table.cell(20, 5).text
                data.append(R20C5)
                R20C6 = table.cell(20, 6).text
                data.append(R20C6)
                R20C7 = table.cell(20, 7).text
                data.append(R20C7)
                R20C8 = table.cell(20, 8).text
                data.append(R20C8)
                R20C9 = table.cell(20, 9).text
                data.append(R20C9)
                R20C10 = table.cell(20, 10).text
                data.append(R20C10)
                R20C12 = table.cell(20, 12).text
                data.append(R20C12)
                R20C14 = table.cell(20, 14).text
                data.append(R20C14)
                R20C15 = table.cell(20, 15).text
                data.append(R20C15)
                R20C16 = table.cell(20, 16).text
                data.append(R20C16)
                R20C17 = table.cell(20, 17).text
                data.append(R20C17)
                R20C18 = table.cell(20, 18).text
                data.append(R20C18)
                R21C1 = table.cell(21, 1).text
                data.append(R21C1)
                R21C2 = table.cell(21, 2).text
                data.append(R21C2)
                R21C3 = table.cell(21, 3).text
                data.append(R21C3)
                R21C4 = table.cell(21, 4).text
                data.append(R21C4)
                R21C5 = table.cell(21, 5).text
                data.append(R21C5)
                R21C6 = table.cell(21, 6).text
                data.append(R21C6)
                R21C7 = table.cell(21, 7).text
                data.append(R21C7)
                R21C8 = table.cell(21, 8).text
                data.append(R21C8)
                R21C9 = table.cell(21, 9).text
                data.append(R21C9)
                R21C10 = table.cell(21, 10).text
                data.append(R21C10)
                R21C12 = table.cell(21, 12).text
                data.append(R21C12)
                R21C14 = table.cell(21, 14).text
                data.append(R21C14)
                R21C15 = table.cell(21, 15).text
                data.append(R21C15)
                R21C16 = table.cell(21, 16).text
                data.append(R21C16)
                R21C17 = table.cell(21, 17).text
                data.append(R21C17)
                R21C18 = table.cell(21, 18).text
                data.append(R21C18)
                R22C1 = table.cell(22, 1).text
                data.append(R22C1)
                R22C10 = table.cell(22, 10).text
                data.append(R22C10)
                R23C10 = table.cell(23, 10).text
                data.append(R23C10)
                R24C0 = table.cell(24, 0).text
                data.append(R24C0)
                R25C0 = table.cell(25, 0).text
                data.append(R25C0)
            elif len_celli == 475:
                R0C0 = ("475"+table.cell(0, 0).text)
                data.append(R0C0)
                R1C0 = table.cell(1, 0).text
                data.append(R1C0)
                R2C1 = table.cell(2, 1).text
                data.append(R2C1)
                R3C1 = table.cell(3, 1).text
                data.append(R3C1)
                R3C13 = table.cell(3, 13).text
                data.append(R3C13)
                R4C1 = table.cell(4, 1).text
                data.append(R4C1)
                R4C9 = table.cell(4, 9).text
                data.append(R4C9)
                R5C9 = table.cell(5, 9).text
                data.append(R5C9)
                R6C1 = table.cell(6, 1).text
                data.append(R6C1)
                R6C7 = table.cell(6, 7).text
                data.append(R6C7)
                R6C10 = table.cell(6, 10).text
                data.append(R6C10)
                R7C1 = table.cell(7, 1).text
                data.append(R7C1)
                R7C7 = table.cell(7, 7).text
                data.append(R7C7)
                R8C1 = table.cell(8, 1).text
                data.append(R8C1)
                R8C7 = table.cell(8, 7).text
                data.append(R8C7)
                R9C1 = table.cell(9, 1).text
                data.append(R9C1)
                R9C7 = table.cell(9, 7).text
                data.append(R9C7)
                R12C1 = table.cell(12, 1).text
                data.append(R12C1)
                R12C2 = table.cell(12, 2).text
                data.append(R12C2)
                R12C3 = table.cell(12, 3).text
                data.append(R12C3)
                R12C4 = table.cell(12, 4).text
                data.append(R12C4)
                R12C5 = table.cell(12, 5).text
                data.append(R12C5)
                R12C6 = table.cell(12, 6).text
                data.append(R12C6)
                R12C7 = table.cell(12, 7).text
                data.append(R12C7)
                R12C8 = table.cell(12, 8).text
                data.append(R12C8)
                R12C9 = table.cell(12, 9).text
                data.append(R12C9)
                R12C10 = table.cell(12, 10).text
                data.append(R12C10)
                R12C12 = table.cell(12, 12).text
                data.append(R12C12)
                R12C14 = table.cell(12, 14).text
                data.append(R12C14)
                R12C15 = table.cell(12, 15).text
                data.append(R12C15)
                R12C16 = table.cell(12, 16).text
                data.append(R12C16)
                R12C17 = table.cell(12, 17).text
                data.append(R12C17)
                R12C18 = table.cell(12, 18).text
                data.append(R12C18)
                R13C1 = table.cell(13, 1).text
                data.append(R13C1)
                R13C2 = table.cell(13, 2).text
                data.append(R13C2)
                R13C3 = table.cell(13, 3).text
                data.append(R13C3)
                R13C4 = table.cell(13, 4).text
                data.append(R13C4)
                R13C5 = table.cell(13, 5).text
                data.append(R13C5)
                R13C6 = table.cell(13, 6).text
                data.append(R13C6)
                R13C7 = table.cell(13, 7).text
                data.append(R13C7)
                R13C8 = table.cell(13, 8).text
                data.append(R13C8)
                R13C9 = table.cell(13, 9).text
                data.append(R13C9)
                R13C10 = table.cell(13, 10).text
                data.append(R13C10)
                R13C12 = table.cell(13, 12).text
                data.append(R13C12)
                R13C14 = table.cell(13, 14).text
                data.append(R13C14)
                R13C15 = table.cell(13, 15).text
                data.append(R13C15)
                R13C16 = table.cell(13, 16).text
                data.append(R13C16)
                R13C17 = table.cell(13, 17).text
                data.append(R13C17)
                R13C18 = table.cell(13, 18).text
                data.append(R13C18)
                R14C1 = table.cell(14, 1).text
                data.append(R14C1)
                R14C2 = table.cell(14, 2).text
                data.append(R14C2)
                R14C3 = table.cell(14, 3).text
                data.append(R14C3)
                R14C4 = table.cell(14, 4).text
                data.append(R14C4)
                R14C5 = table.cell(14, 5).text
                data.append(R14C5)
                R14C6 = table.cell(14, 6).text
                data.append(R14C6)
                R14C7 = table.cell(14, 7).text
                data.append(R14C7)
                R14C8 = table.cell(14, 8).text
                data.append(R14C8)
                R14C9 = table.cell(14, 9).text
                data.append(R14C9)
                R14C10 = table.cell(14, 10).text
                data.append(R14C10)
                R14C12 = table.cell(14, 12).text
                data.append(R14C12)
                R14C14 = table.cell(14, 14).text
                data.append(R14C14)
                R14C15 = table.cell(14, 15).text
                data.append(R14C15)
                R14C16 = table.cell(14, 16).text
                data.append(R14C16)
                R14C17 = table.cell(14, 17).text
                data.append(R14C17)
                R14C18 = table.cell(14, 18).text
                data.append(R14C18)
                R15C1 = table.cell(15, 1).text
                data.append(R15C1)
                R15C2 = table.cell(15, 2).text
                data.append(R15C2)
                R15C3 = table.cell(15, 3).text
                data.append(R15C3)
                R15C4 = table.cell(15, 4).text
                data.append(R15C4)
                R15C5 = table.cell(15, 5).text
                data.append(R15C5)
                R15C6 = table.cell(15, 6).text
                data.append(R15C6)
                R15C7 = table.cell(15, 7).text
                data.append(R15C7)
                R15C8 = table.cell(15, 8).text
                data.append(R15C8)
                R15C9 = table.cell(15, 9).text
                data.append(R15C9)
                R15C10 = table.cell(15, 10).text
                data.append(R15C10)
                R15C12 = table.cell(15, 12).text
                data.append(R15C12)
                R15C14 = table.cell(15, 14).text
                data.append(R15C14)
                R15C15 = table.cell(15, 15).text
                data.append(R15C15)
                R15C16 = table.cell(15, 16).text
                data.append(R15C16)
                R15C17 = table.cell(15, 17).text
                data.append(R15C17)
                R15C18 = table.cell(15, 18).text
                data.append(R15C18)
                R16C1 = table.cell(16, 1).text
                data.append(R16C1)
                R16C2 = table.cell(16, 2).text
                data.append(R16C2)
                R16C3 = table.cell(16, 3).text
                data.append(R16C3)
                R16C4 = table.cell(16, 4).text
                data.append(R16C4)
                R16C5 = table.cell(16, 5).text
                data.append(R16C5)
                R16C6 = table.cell(16, 6).text
                data.append(R16C6)
                R16C7 = table.cell(16, 7).text
                data.append(R16C7)
                R16C8 = table.cell(16, 8).text
                data.append(R16C8)
                R16C9 = table.cell(16, 9).text
                data.append(R16C9)
                R16C10 = table.cell(16, 10).text
                data.append(R16C10)
                R16C12 = table.cell(16, 12).text
                data.append(R16C12)
                R16C14 = table.cell(16, 14).text
                data.append(R16C14)
                R16C15 = table.cell(16, 15).text
                data.append(R16C15)
                R16C16 = table.cell(16, 16).text
                data.append(R16C16)
                R16C17 = table.cell(16, 17).text
                data.append(R16C17)
                R16C18 = table.cell(16, 18).text
                data.append(R16C18)
                R17C1 = table.cell(17, 1).text
                data.append(R17C1)
                R17C2 = table.cell(17, 2).text
                data.append(R17C2)
                R17C3 = table.cell(17, 3).text
                data.append(R17C3)
                R17C4 = table.cell(17, 4).text
                data.append(R17C4)
                R17C5 = table.cell(17, 5).text
                data.append(R17C5)
                R17C6 = table.cell(17, 6).text
                data.append(R17C6)
                R17C7 = table.cell(17, 7).text
                data.append(R17C7)
                R17C8 = table.cell(17, 8).text
                data.append(R17C8)
                R17C9 = table.cell(17, 9).text
                data.append(R17C9)
                R17C10 = table.cell(17, 10).text
                data.append(R17C10)
                R17C12 = table.cell(17, 12).text
                data.append(R17C12)
                R17C14 = table.cell(17, 14).text
                data.append(R17C14)
                R17C15 = table.cell(17, 15).text
                data.append(R17C15)
                R17C16 = table.cell(17, 16).text
                data.append(R17C16)
                R17C17 = table.cell(17, 17).text
                data.append(R17C17)
                R17C18 = table.cell(17, 18).text
                data.append(R17C18)
                R18C1 = table.cell(18, 1).text
                data.append(R18C1)
                R18C2 = table.cell(18, 2).text
                data.append(R18C2)
                R18C3 = table.cell(18, 3).text
                data.append(R18C3)
                R18C4 = table.cell(18, 4).text
                data.append(R18C4)
                R18C5 = table.cell(18, 5).text
                data.append(R18C5)
                R18C6 = table.cell(18, 6).text
                data.append(R18C6)
                R18C7 = table.cell(18, 7).text
                data.append(R18C7)
                R18C8 = table.cell(18, 8).text
                data.append(R18C8)
                R18C9 = table.cell(18, 9).text
                data.append(R18C9)
                R18C10 = table.cell(18, 10).text
                data.append(R18C10)
                R18C12 = table.cell(18, 12).text
                data.append(R18C12)
                R18C14 = table.cell(18, 14).text
                data.append(R18C14)
                R18C15 = table.cell(18, 15).text
                data.append(R18C15)
                R18C16 = table.cell(18, 16).text
                data.append(R18C16)
                R18C17 = table.cell(18, 17).text
                data.append(R18C17)
                R18C18 = table.cell(18, 18).text
                data.append(R18C18)
                R19C1 = table.cell(19, 1).text
                data.append(R19C1)
                R19C2 = table.cell(19, 2).text
                data.append(R19C2)
                R19C3 = table.cell(19, 3).text
                data.append(R19C3)
                R19C4 = table.cell(19, 4).text
                data.append(R19C4)
                R19C5 = table.cell(19, 5).text
                data.append(R19C5)
                R19C6 = table.cell(19, 6).text
                data.append(R19C6)
                R19C7 = table.cell(19, 7).text
                data.append(R19C7)
                R19C8 = table.cell(19, 8).text
                data.append(R19C8)
                R19C9 = table.cell(19, 9).text
                data.append(R19C9)
                R19C10 = table.cell(19, 10).text
                data.append(R19C10)
                R19C12 = table.cell(19, 12).text
                data.append(R19C12)
                R19C14 = table.cell(19, 14).text
                data.append(R19C14)
                R19C15 = table.cell(19, 15).text
                data.append(R19C15)
                R19C16 = table.cell(19, 16).text
                data.append(R19C16)
                R19C17 = table.cell(19, 17).text
                data.append(R19C17)
                R19C18 = table.cell(19, 18).text
                data.append(R19C18)
                R20C1 = table.cell(20, 1).text
                data.append(R20C1)
                R20C2 = table.cell(20, 2).text
                data.append(R20C2)
                R20C3 = table.cell(20, 3).text
                data.append(R20C3)
                R20C4 = table.cell(20, 4).text
                data.append(R20C4)
                R20C5 = table.cell(20, 5).text
                data.append(R20C5)
                R20C6 = table.cell(20, 6).text
                data.append(R20C6)
                R20C7 = table.cell(20, 7).text
                data.append(R20C7)
                R20C8 = table.cell(20, 8).text
                data.append(R20C8)
                R20C9 = table.cell(20, 9).text
                data.append(R20C9)
                R20C10 = table.cell(20, 10).text
                data.append(R20C10)
                R20C12 = table.cell(20, 12).text
                data.append(R20C12)
                R20C14 = table.cell(20, 14).text
                data.append(R20C14)
                R20C15 = table.cell(20, 15).text
                data.append(R20C15)
                R20C16 = table.cell(20, 16).text
                data.append(R20C16)
                R20C17 = table.cell(20, 17).text
                data.append(R20C17)
                R20C18 = table.cell(20, 18).text
                data.append(R20C18)
                R21C1 = table.cell(21, 1).text
                data.append(R21C1)
                R21C10 = table.cell(21, 10).text
                data.append(R21C10)
                R22C10 = table.cell(22, 10).text
                data.append(R22C10)
                R23C0 = table.cell(23, 0).text
                data.append(R23C0)
                R24C0 = table.cell(24, 0).text
                data.append(R24C0)
            elif len_celli == 456:
                R0C0 = ("456"+table.cell(0, 0).text)
                data.append(R0C0)
                R1C0 = table.cell(1, 0).text
                data.append(R1C0)
                R2C1 = table.cell(2, 1).text
                data.append(R2C1)
                R3C1 = table.cell(3, 1).text
                data.append(R3C1)
                R3C13 = table.cell(3, 13).text
                data.append(R3C13)
                R4C1 = table.cell(4, 1).text
                data.append(R4C1)
                R4C9 = table.cell(4, 9).text
                data.append(R4C9)
                R5C9 = table.cell(5, 9).text
                data.append(R5C9)
                R6C1 = table.cell(6, 1).text
                data.append(R6C1)
                R6C7 = table.cell(6, 7).text
                data.append(R6C7)
                R6C10 = table.cell(6, 10).text
                data.append(R6C10)
                R7C1 = table.cell(7, 1).text
                data.append(R7C1)
                R7C7 = table.cell(7, 7).text
                data.append(R7C7)
                R8C1 = table.cell(8, 1).text
                data.append(R8C1)
                R8C7 = table.cell(8, 7).text
                data.append(R8C7)
                R9C1 = table.cell(9, 1).text
                data.append(R9C1)
                R9C7 = table.cell(9, 7).text
                data.append(R9C7)
                R12C1 = table.cell(12, 1).text
                data.append(R12C1)
                R12C2 = table.cell(12, 2).text
                data.append(R12C2)
                R12C3 = table.cell(12, 3).text
                data.append(R12C3)
                R12C4 = table.cell(12, 4).text
                data.append(R12C4)
                R12C5 = table.cell(12, 5).text
                data.append(R12C5)
                R12C6 = table.cell(12, 6).text
                data.append(R12C6)
                R12C7 = table.cell(12, 7).text
                data.append(R12C7)
                R12C8 = table.cell(12, 8).text
                data.append(R12C8)
                R12C9 = table.cell(12, 9).text
                data.append(R12C9)
                R12C10 = table.cell(12, 10).text
                data.append(R12C10)
                R12C12 = table.cell(12, 12).text
                data.append(R12C12)
                R12C14 = table.cell(12, 14).text
                data.append(R12C14)
                R12C15 = table.cell(12, 15).text
                data.append(R12C15)
                R12C16 = table.cell(12, 16).text
                data.append(R12C16)
                R12C17 = table.cell(12, 17).text
                data.append(R12C17)
                R12C18 = table.cell(12, 18).text
                data.append(R12C18)
                R13C1 = table.cell(13, 1).text
                data.append(R13C1)
                R13C2 = table.cell(13, 2).text
                data.append(R13C2)
                R13C3 = table.cell(13, 3).text
                data.append(R13C3)
                R13C4 = table.cell(13, 4).text
                data.append(R13C4)
                R13C5 = table.cell(13, 5).text
                data.append(R13C5)
                R13C6 = table.cell(13, 6).text
                data.append(R13C6)
                R13C7 = table.cell(13, 7).text
                data.append(R13C7)
                R13C8 = table.cell(13, 8).text
                data.append(R13C8)
                R13C9 = table.cell(13, 9).text
                data.append(R13C9)
                R13C10 = table.cell(13, 10).text
                data.append(R13C10)
                R13C12 = table.cell(13, 12).text
                data.append(R13C12)
                R13C14 = table.cell(13, 14).text
                data.append(R13C14)
                R13C15 = table.cell(13, 15).text
                data.append(R13C15)
                R13C16 = table.cell(13, 16).text
                data.append(R13C16)
                R13C17 = table.cell(13, 17).text
                data.append(R13C17)
                R13C18 = table.cell(13, 18).text
                data.append(R13C18)
                R14C1 = table.cell(14, 1).text
                data.append(R14C1)
                R14C2 = table.cell(14, 2).text
                data.append(R14C2)
                R14C3 = table.cell(14, 3).text
                data.append(R14C3)
                R14C4 = table.cell(14, 4).text
                data.append(R14C4)
                R14C5 = table.cell(14, 5).text
                data.append(R14C5)
                R14C6 = table.cell(14, 6).text
                data.append(R14C6)
                R14C7 = table.cell(14, 7).text
                data.append(R14C7)
                R14C8 = table.cell(14, 8).text
                data.append(R14C8)
                R14C9 = table.cell(14, 9).text
                data.append(R14C9)
                R14C10 = table.cell(14, 10).text
                data.append(R14C10)
                R14C12 = table.cell(14, 12).text
                data.append(R14C12)
                R14C14 = table.cell(14, 14).text
                data.append(R14C14)
                R14C15 = table.cell(14, 15).text
                data.append(R14C15)
                R14C16 = table.cell(14, 16).text
                data.append(R14C16)
                R14C17 = table.cell(14, 17).text
                data.append(R14C17)
                R14C18 = table.cell(14, 18).text
                data.append(R14C18)
                R15C1 = table.cell(15, 1).text
                data.append(R15C1)
                R15C2 = table.cell(15, 2).text
                data.append(R15C2)
                R15C3 = table.cell(15, 3).text
                data.append(R15C3)
                R15C4 = table.cell(15, 4).text
                data.append(R15C4)
                R15C5 = table.cell(15, 5).text
                data.append(R15C5)
                R15C6 = table.cell(15, 6).text
                data.append(R15C6)
                R15C7 = table.cell(15, 7).text
                data.append(R15C7)
                R15C8 = table.cell(15, 8).text
                data.append(R15C8)
                R15C9 = table.cell(15, 9).text
                data.append(R15C9)
                R15C10 = table.cell(15, 10).text
                data.append(R15C10)
                R15C12 = table.cell(15, 12).text
                data.append(R15C12)
                R15C14 = table.cell(15, 14).text
                data.append(R15C14)
                R15C15 = table.cell(15, 15).text
                data.append(R15C15)
                R15C16 = table.cell(15, 16).text
                data.append(R15C16)
                R15C17 = table.cell(15, 17).text
                data.append(R15C17)
                R15C18 = table.cell(15, 18).text
                data.append(R15C18)
                R16C1 = table.cell(16, 1).text
                data.append(R16C1)
                R16C2 = table.cell(16, 2).text
                data.append(R16C2)
                R16C3 = table.cell(16, 3).text
                data.append(R16C3)
                R16C4 = table.cell(16, 4).text
                data.append(R16C4)
                R16C5 = table.cell(16, 5).text
                data.append(R16C5)
                R16C6 = table.cell(16, 6).text
                data.append(R16C6)
                R16C7 = table.cell(16, 7).text
                data.append(R16C7)
                R16C8 = table.cell(16, 8).text
                data.append(R16C8)
                R16C9 = table.cell(16, 9).text
                data.append(R16C9)
                R16C10 = table.cell(16, 10).text
                data.append(R16C10)
                R16C12 = table.cell(16, 12).text
                data.append(R16C12)
                R16C14 = table.cell(16, 14).text
                data.append(R16C14)
                R16C15 = table.cell(16, 15).text
                data.append(R16C15)
                R16C16 = table.cell(16, 16).text
                data.append(R16C16)
                R16C17 = table.cell(16, 17).text
                data.append(R16C17)
                R16C18 = table.cell(16, 18).text
                data.append(R16C18)
                R17C1 = table.cell(17, 1).text
                data.append(R17C1)
                R17C2 = table.cell(17, 2).text
                data.append(R17C2)
                R17C3 = table.cell(17, 3).text
                data.append(R17C3)
                R17C4 = table.cell(17, 4).text
                data.append(R17C4)
                R17C5 = table.cell(17, 5).text
                data.append(R17C5)
                R17C6 = table.cell(17, 6).text
                data.append(R17C6)
                R17C7 = table.cell(17, 7).text
                data.append(R17C7)
                R17C8 = table.cell(17, 8).text
                data.append(R17C8)
                R17C9 = table.cell(17, 9).text
                data.append(R17C9)
                R17C10 = table.cell(17, 10).text
                data.append(R17C10)
                R17C12 = table.cell(17, 12).text
                data.append(R17C12)
                R17C14 = table.cell(17, 14).text
                data.append(R17C14)
                R17C15 = table.cell(17, 15).text
                data.append(R17C15)
                R17C16 = table.cell(17, 16).text
                data.append(R17C16)
                R17C17 = table.cell(17, 17).text
                data.append(R17C17)
                R17C18 = table.cell(17, 18).text
                data.append(R17C18)
                R18C1 = table.cell(18, 1).text
                data.append(R18C1)
                R18C2 = table.cell(18, 2).text
                data.append(R18C2)
                R18C3 = table.cell(18, 3).text
                data.append(R18C3)
                R18C4 = table.cell(18, 4).text
                data.append(R18C4)
                R18C5 = table.cell(18, 5).text
                data.append(R18C5)
                R18C6 = table.cell(18, 6).text
                data.append(R18C6)
                R18C7 = table.cell(18, 7).text
                data.append(R18C7)
                R18C8 = table.cell(18, 8).text
                data.append(R18C8)
                R18C9 = table.cell(18, 9).text
                data.append(R18C9)
                R18C10 = table.cell(18, 10).text
                data.append(R18C10)
                R18C12 = table.cell(18, 12).text
                data.append(R18C12)
                R18C14 = table.cell(18, 14).text
                data.append(R18C14)
                R18C15 = table.cell(18, 15).text
                data.append(R18C15)
                R18C16 = table.cell(18, 16).text
                data.append(R18C16)
                R18C17 = table.cell(18, 17).text
                data.append(R18C17)
                R18C18 = table.cell(18, 18).text
                data.append(R18C18)
                R19C1 = table.cell(19, 1).text
                data.append(R19C1)
                R19C2 = table.cell(19, 2).text
                data.append(R19C2)
                R19C3 = table.cell(19, 3).text
                data.append(R19C3)
                R19C4 = table.cell(19, 4).text
                data.append(R19C4)
                R19C5 = table.cell(19, 5).text
                data.append(R19C5)
                R19C6 = table.cell(19, 6).text
                data.append(R19C6)
                R19C7 = table.cell(19, 7).text
                data.append(R19C7)
                R19C8 = table.cell(19, 8).text
                data.append(R19C8)
                R19C9 = table.cell(19, 9).text
                data.append(R19C9)
                R19C10 = table.cell(19, 10).text
                data.append(R19C10)
                R19C12 = table.cell(19, 12).text
                data.append(R19C12)
                R19C14 = table.cell(19, 14).text
                data.append(R19C14)
                R19C15 = table.cell(19, 15).text
                data.append(R19C15)
                R19C16 = table.cell(19, 16).text
                data.append(R19C16)
                R19C17 = table.cell(19, 17).text
                data.append(R19C17)
                R19C18 = table.cell(19, 18).text
                data.append(R19C18)
                R20C1 = table.cell(20, 1).text
                data.append(R20C1)
                R20C10 = table.cell(20, 10).text
                data.append(R20C10)
                R21C10 = table.cell(21, 10).text
                data.append(R21C10)
                R22C0 = table.cell(22, 0).text
                data.append(R22C0)
                R23C0 = table.cell(23, 0).text
                data.append(R23C0)
            elif len_celli == 437:
                R0C0 = ("437"+table.cell(0, 0).text)
                data.append(R0C0)
                R1C0 = table.cell(1, 0).text
                data.append(R1C0)
                R2C1 = table.cell(2, 1).text
                data.append(R2C1)
                R3C1 = table.cell(3, 1).text
                data.append(R3C1)
                R3C13 = table.cell(3, 13).text
                data.append(R3C13)
                R4C1 = table.cell(4, 1).text
                data.append(R4C1)
                R4C9 = table.cell(4, 9).text
                data.append(R4C9)
                R5C9 = table.cell(5, 9).text
                data.append(R5C9)
                R6C1 = table.cell(6, 1).text
                data.append(R6C1)
                R6C7 = table.cell(6, 7).text
                data.append(R6C7)
                R6C10 = table.cell(6, 10).text
                data.append(R6C10)
                R7C1 = table.cell(7, 1).text
                data.append(R7C1)
                R7C7 = table.cell(7, 7).text
                data.append(R7C7)
                R8C1 = table.cell(8, 1).text
                data.append(R8C1)
                R8C7 = table.cell(8, 7).text
                data.append(R8C7)
                R9C1 = table.cell(9, 1).text
                data.append(R9C1)
                R9C7 = table.cell(9, 7).text
                data.append(R9C7)
                R12C1 = table.cell(12, 1).text
                data.append(R12C1)
                R12C2 = table.cell(12, 2).text
                data.append(R12C2)
                R12C3 = table.cell(12, 3).text
                data.append(R12C3)
                R12C4 = table.cell(12, 4).text
                data.append(R12C4)
                R12C5 = table.cell(12, 5).text
                data.append(R12C5)
                R12C6 = table.cell(12, 6).text
                data.append(R12C6)
                R12C7 = table.cell(12, 7).text
                data.append(R12C7)
                R12C8 = table.cell(12, 8).text
                data.append(R12C8)
                R12C9 = table.cell(12, 9).text
                data.append(R12C9)
                R12C10 = table.cell(12, 10).text
                data.append(R12C10)
                R12C12 = table.cell(12, 12).text
                data.append(R12C12)
                R12C14 = table.cell(12, 14).text
                data.append(R12C14)
                R12C15 = table.cell(12, 15).text
                data.append(R12C15)
                R12C16 = table.cell(12, 16).text
                data.append(R12C16)
                R12C17 = table.cell(12, 17).text
                data.append(R12C17)
                R12C18 = table.cell(12, 18).text
                data.append(R12C18)
                R13C1 = table.cell(13, 1).text
                data.append(R13C1)
                R13C2 = table.cell(13, 2).text
                data.append(R13C2)
                R13C3 = table.cell(13, 3).text
                data.append(R13C3)
                R13C4 = table.cell(13, 4).text
                data.append(R13C4)
                R13C5 = table.cell(13, 5).text
                data.append(R13C5)
                R13C6 = table.cell(13, 6).text
                data.append(R13C6)
                R13C7 = table.cell(13, 7).text
                data.append(R13C7)
                R13C8 = table.cell(13, 8).text
                data.append(R13C8)
                R13C9 = table.cell(13, 9).text
                data.append(R13C9)
                R13C10 = table.cell(13, 10).text
                data.append(R13C10)
                R13C12 = table.cell(13, 12).text
                data.append(R13C12)
                R13C14 = table.cell(13, 14).text
                data.append(R13C14)
                R13C15 = table.cell(13, 15).text
                data.append(R13C15)
                R13C16 = table.cell(13, 16).text
                data.append(R13C16)
                R13C17 = table.cell(13, 17).text
                data.append(R13C17)
                R13C18 = table.cell(13, 18).text
                data.append(R13C18)
                R14C1 = table.cell(14, 1).text
                data.append(R14C1)
                R14C2 = table.cell(14, 2).text
                data.append(R14C2)
                R14C3 = table.cell(14, 3).text
                data.append(R14C3)
                R14C4 = table.cell(14, 4).text
                data.append(R14C4)
                R14C5 = table.cell(14, 5).text
                data.append(R14C5)
                R14C6 = table.cell(14, 6).text
                data.append(R14C6)
                R14C7 = table.cell(14, 7).text
                data.append(R14C7)
                R14C8 = table.cell(14, 8).text
                data.append(R14C8)
                R14C9 = table.cell(14, 9).text
                data.append(R14C9)
                R14C10 = table.cell(14, 10).text
                data.append(R14C10)
                R14C12 = table.cell(14, 12).text
                data.append(R14C12)
                R14C14 = table.cell(14, 14).text
                data.append(R14C14)
                R14C15 = table.cell(14, 15).text
                data.append(R14C15)
                R14C16 = table.cell(14, 16).text
                data.append(R14C16)
                R14C17 = table.cell(14, 17).text
                data.append(R14C17)
                R14C18 = table.cell(14, 18).text
                data.append(R14C18)
                R15C1 = table.cell(15, 1).text
                data.append(R15C1)
                R15C2 = table.cell(15, 2).text
                data.append(R15C2)
                R15C3 = table.cell(15, 3).text
                data.append(R15C3)
                R15C4 = table.cell(15, 4).text
                data.append(R15C4)
                R15C5 = table.cell(15, 5).text
                data.append(R15C5)
                R15C6 = table.cell(15, 6).text
                data.append(R15C6)
                R15C7 = table.cell(15, 7).text
                data.append(R15C7)
                R15C8 = table.cell(15, 8).text
                data.append(R15C8)
                R15C9 = table.cell(15, 9).text
                data.append(R15C9)
                R15C10 = table.cell(15, 10).text
                data.append(R15C10)
                R15C12 = table.cell(15, 12).text
                data.append(R15C12)
                R15C14 = table.cell(15, 14).text
                data.append(R15C14)
                R15C15 = table.cell(15, 15).text
                data.append(R15C15)
                R15C16 = table.cell(15, 16).text
                data.append(R15C16)
                R15C17 = table.cell(15, 17).text
                data.append(R15C17)
                R15C18 = table.cell(15, 18).text
                data.append(R15C18)
                R16C1 = table.cell(16, 1).text
                data.append(R16C1)
                R16C2 = table.cell(16, 2).text
                data.append(R16C2)
                R16C3 = table.cell(16, 3).text
                data.append(R16C3)
                R16C4 = table.cell(16, 4).text
                data.append(R16C4)
                R16C5 = table.cell(16, 5).text
                data.append(R16C5)
                R16C6 = table.cell(16, 6).text
                data.append(R16C6)
                R16C7 = table.cell(16, 7).text
                data.append(R16C7)
                R16C8 = table.cell(16, 8).text
                data.append(R16C8)
                R16C9 = table.cell(16, 9).text
                data.append(R16C9)
                R16C10 = table.cell(16, 10).text
                data.append(R16C10)
                R16C12 = table.cell(16, 12).text
                data.append(R16C12)
                R16C14 = table.cell(16, 14).text
                data.append(R16C14)
                R16C15 = table.cell(16, 15).text
                data.append(R16C15)
                R16C16 = table.cell(16, 16).text
                data.append(R16C16)
                R16C17 = table.cell(16, 17).text
                data.append(R16C17)
                R16C18 = table.cell(16, 18).text
                data.append(R16C18)
                R17C1 = table.cell(17, 1).text
                data.append(R17C1)
                R17C2 = table.cell(17, 2).text
                data.append(R17C2)
                R17C3 = table.cell(17, 3).text
                data.append(R17C3)
                R17C4 = table.cell(17, 4).text
                data.append(R17C4)
                R17C5 = table.cell(17, 5).text
                data.append(R17C5)
                R17C6 = table.cell(17, 6).text
                data.append(R17C6)
                R17C7 = table.cell(17, 7).text
                data.append(R17C7)
                R17C8 = table.cell(17, 8).text
                data.append(R17C8)
                R17C9 = table.cell(17, 9).text
                data.append(R17C9)
                R17C10 = table.cell(17, 10).text
                data.append(R17C10)
                R17C12 = table.cell(17, 12).text
                data.append(R17C12)
                R17C14 = table.cell(17, 14).text
                data.append(R17C14)
                R17C15 = table.cell(17, 15).text
                data.append(R17C15)
                R17C16 = table.cell(17, 16).text
                data.append(R17C16)
                R17C17 = table.cell(17, 17).text
                data.append(R17C17)
                R17C18 = table.cell(17, 18).text
                data.append(R17C18)
                R18C1 = table.cell(18, 1).text
                data.append(R18C1)
                R18C2 = table.cell(18, 2).text
                data.append(R18C2)
                R18C3 = table.cell(18, 3).text
                data.append(R18C3)
                R18C4 = table.cell(18, 4).text
                data.append(R18C4)
                R18C5 = table.cell(18, 5).text
                data.append(R18C5)
                R18C6 = table.cell(18, 6).text
                data.append(R18C6)
                R18C7 = table.cell(18, 7).text
                data.append(R18C7)
                R18C8 = table.cell(18, 8).text
                data.append(R18C8)
                R18C9 = table.cell(18, 9).text
                data.append(R18C9)
                R18C10 = table.cell(18, 10).text
                data.append(R18C10)
                R18C12 = table.cell(18, 12).text
                data.append(R18C12)
                R18C14 = table.cell(18, 14).text
                data.append(R18C14)
                R18C15 = table.cell(18, 15).text
                data.append(R18C15)
                R18C16 = table.cell(18, 16).text
                data.append(R18C16)
                R18C17 = table.cell(18, 17).text
                data.append(R18C17)
                R18C18 = table.cell(18, 18).text
                data.append(R18C18)
                R19C1 = table.cell(19, 1).text
                data.append(R19C1)
                R19C10 = table.cell(19, 10).text
                data.append(R19C10)
                R20C10 = table.cell(20, 10).text
                data.append(R20C10)
                R21C0 = table.cell(21, 0).text
                data.append(R21C0)
                R22C0 = table.cell(22, 0).text
                data.append(R22C0)
            elif len_celli == 418:
                R0C0 = ("418"+table.cell(0, 0).text)
                data.append(R0C0)
                R1C0 = table.cell(1, 0).text
                data.append(R1C0)
                R2C1 = table.cell(2, 1).text
                data.append(R2C1)
                R3C1 = table.cell(3, 1).text
                data.append(R3C1)
                R3C13 = table.cell(3, 13).text
                data.append(R3C13)
                R4C1 = table.cell(4, 1).text
                data.append(R4C1)
                R4C9 = table.cell(4, 9).text
                data.append(R4C9)
                R5C9 = table.cell(5, 9).text
                data.append(R5C9)
                R6C1 = table.cell(6, 1).text
                data.append(R6C1)
                R6C7 = table.cell(6, 7).text
                data.append(R6C7)
                R6C10 = table.cell(6, 10).text
                data.append(R6C10)
                R7C1 = table.cell(7, 1).text
                data.append(R7C1)
                R7C7 = table.cell(7, 7).text
                data.append(R7C7)
                R8C1 = table.cell(8, 1).text
                data.append(R8C1)
                R8C7 = table.cell(8, 7).text
                data.append(R8C7)
                R9C1 = table.cell(9, 1).text
                data.append(R9C1)
                R9C7 = table.cell(9, 7).text
                data.append(R9C7)
                R12C1 = table.cell(12, 1).text
                data.append(R12C1)
                R12C2 = table.cell(12, 2).text
                data.append(R12C2)
                R12C3 = table.cell(12, 3).text
                data.append(R12C3)
                R12C4 = table.cell(12, 4).text
                data.append(R12C4)
                R12C5 = table.cell(12, 5).text
                data.append(R12C5)
                R12C6 = table.cell(12, 6).text
                data.append(R12C6)
                R12C7 = table.cell(12, 7).text
                data.append(R12C7)
                R12C8 = table.cell(12, 8).text
                data.append(R12C8)
                R12C9 = table.cell(12, 9).text
                data.append(R12C9)
                R12C10 = table.cell(12, 10).text
                data.append(R12C10)
                R12C12 = table.cell(12, 12).text
                data.append(R12C12)
                R12C14 = table.cell(12, 14).text
                data.append(R12C14)
                R12C15 = table.cell(12, 15).text
                data.append(R12C15)
                R12C16 = table.cell(12, 16).text
                data.append(R12C16)
                R12C17 = table.cell(12, 17).text
                data.append(R12C17)
                R12C18 = table.cell(12, 18).text
                data.append(R12C18)
                R13C1 = table.cell(13, 1).text
                data.append(R13C1)
                R13C2 = table.cell(13, 2).text
                data.append(R13C2)
                R13C3 = table.cell(13, 3).text
                data.append(R13C3)
                R13C4 = table.cell(13, 4).text
                data.append(R13C4)
                R13C5 = table.cell(13, 5).text
                data.append(R13C5)
                R13C6 = table.cell(13, 6).text
                data.append(R13C6)
                R13C7 = table.cell(13, 7).text
                data.append(R13C7)
                R13C8 = table.cell(13, 8).text
                data.append(R13C8)
                R13C9 = table.cell(13, 9).text
                data.append(R13C9)
                R13C10 = table.cell(13, 10).text
                data.append(R13C10)
                R13C12 = table.cell(13, 12).text
                data.append(R13C12)
                R13C14 = table.cell(13, 14).text
                data.append(R13C14)
                R13C15 = table.cell(13, 15).text
                data.append(R13C15)
                R13C16 = table.cell(13, 16).text
                data.append(R13C16)
                R13C17 = table.cell(13, 17).text
                data.append(R13C17)
                R13C18 = table.cell(13, 18).text
                data.append(R13C18)
                R14C1 = table.cell(14, 1).text
                data.append(R14C1)
                R14C2 = table.cell(14, 2).text
                data.append(R14C2)
                R14C3 = table.cell(14, 3).text
                data.append(R14C3)
                R14C4 = table.cell(14, 4).text
                data.append(R14C4)
                R14C5 = table.cell(14, 5).text
                data.append(R14C5)
                R14C6 = table.cell(14, 6).text
                data.append(R14C6)
                R14C7 = table.cell(14, 7).text
                data.append(R14C7)
                R14C8 = table.cell(14, 8).text
                data.append(R14C8)
                R14C9 = table.cell(14, 9).text
                data.append(R14C9)
                R14C10 = table.cell(14, 10).text
                data.append(R14C10)
                R14C12 = table.cell(14, 12).text
                data.append(R14C12)
                R14C14 = table.cell(14, 14).text
                data.append(R14C14)
                R14C15 = table.cell(14, 15).text
                data.append(R14C15)
                R14C16 = table.cell(14, 16).text
                data.append(R14C16)
                R14C17 = table.cell(14, 17).text
                data.append(R14C17)
                R14C18 = table.cell(14, 18).text
                data.append(R14C18)
                R15C1 = table.cell(15, 1).text
                data.append(R15C1)
                R15C2 = table.cell(15, 2).text
                data.append(R15C2)
                R15C3 = table.cell(15, 3).text
                data.append(R15C3)
                R15C4 = table.cell(15, 4).text
                data.append(R15C4)
                R15C5 = table.cell(15, 5).text
                data.append(R15C5)
                R15C6 = table.cell(15, 6).text
                data.append(R15C6)
                R15C7 = table.cell(15, 7).text
                data.append(R15C7)
                R15C8 = table.cell(15, 8).text
                data.append(R15C8)
                R15C9 = table.cell(15, 9).text
                data.append(R15C9)
                R15C10 = table.cell(15, 10).text
                data.append(R15C10)
                R15C12 = table.cell(15, 12).text
                data.append(R15C12)
                R15C14 = table.cell(15, 14).text
                data.append(R15C14)
                R15C15 = table.cell(15, 15).text
                data.append(R15C15)
                R15C16 = table.cell(15, 16).text
                data.append(R15C16)
                R15C17 = table.cell(15, 17).text
                data.append(R15C17)
                R15C18 = table.cell(15, 18).text
                data.append(R15C18)
                R16C1 = table.cell(16, 1).text
                data.append(R16C1)
                R16C2 = table.cell(16, 2).text
                data.append(R16C2)
                R16C3 = table.cell(16, 3).text
                data.append(R16C3)
                R16C4 = table.cell(16, 4).text
                data.append(R16C4)
                R16C5 = table.cell(16, 5).text
                data.append(R16C5)
                R16C6 = table.cell(16, 6).text
                data.append(R16C6)
                R16C7 = table.cell(16, 7).text
                data.append(R16C7)
                R16C8 = table.cell(16, 8).text
                data.append(R16C8)
                R16C9 = table.cell(16, 9).text
                data.append(R16C9)
                R16C10 = table.cell(16, 10).text
                data.append(R16C10)
                R16C12 = table.cell(16, 12).text
                data.append(R16C12)
                R16C14 = table.cell(16, 14).text
                data.append(R16C14)
                R16C15 = table.cell(16, 15).text
                data.append(R16C15)
                R16C16 = table.cell(16, 16).text
                data.append(R16C16)
                R16C17 = table.cell(16, 17).text
                data.append(R16C17)
                R16C18 = table.cell(16, 18).text
                data.append(R16C18)
                R17C1 = table.cell(17, 1).text
                data.append(R17C1)
                R17C2 = table.cell(17, 2).text
                data.append(R17C2)
                R17C3 = table.cell(17, 3).text
                data.append(R17C3)
                R17C4 = table.cell(17, 4).text
                data.append(R17C4)
                R17C5 = table.cell(17, 5).text
                data.append(R17C5)
                R17C6 = table.cell(17, 6).text
                data.append(R17C6)
                R17C7 = table.cell(17, 7).text
                data.append(R17C7)
                R17C8 = table.cell(17, 8).text
                data.append(R17C8)
                R17C9 = table.cell(17, 9).text
                data.append(R17C9)
                R17C10 = table.cell(17, 10).text
                data.append(R17C10)
                R17C12 = table.cell(17, 12).text
                data.append(R17C12)
                R17C14 = table.cell(17, 14).text
                data.append(R17C14)
                R17C15 = table.cell(17, 15).text
                data.append(R17C15)
                R17C16 = table.cell(17, 16).text
                data.append(R17C16)
                R17C17 = table.cell(17, 17).text
                data.append(R17C17)
                R17C18 = table.cell(17, 18).text
                data.append(R17C18)
                R18C1 = table.cell(18, 1).text
                data.append(R18C1)
                R18C10 = table.cell(18, 10).text
                data.append(R18C10)
                R19C10 = table.cell(19, 10).text
                data.append(R19C10)
                R20C0 = table.cell(20, 0).text
                data.append(R20C0)
                R21C0 = table.cell(21, 0).text
                data.append(R21C0)
            else:
                print "wrong"
            l_ = range(len(data))
            x = data
            for j in l_:
                sheet1.write(i + 1, j, x[j])  # 第一个是写入哪一行,第二个写入参数的列,第三个是要写入的数据
            print (str(i+1)+'_'+ListName[i], "写入完成")
    
#            t.append(x)
#            print (str(len(t-1))+'_'+''.join(t[i-1]))
#        if i>9:
#           break
    f.save((rootdir+"数据" + nowtime + ".xls").decode(encoding="UTF-8", errors="strict"))  # 保存文件
print "完成数据"+str(i+1)+"条"
