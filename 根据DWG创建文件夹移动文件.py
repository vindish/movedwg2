#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import shutil

# 读取中文路径 u''
rootpath = unicode((os.getcwd()).replace("\\","/"),'gb2312')
dirs = os.listdir(rootpath)

def mkdir(path):
    folder = os.path.exists(path)
    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)  # makedirs 创建文件时如果路径不存在会创建这个路径
    else:
        print "---  There is this folder!  ---"

filelista = []
for dirfile1 in dirs:
    fileAllPath = os.path.join(rootpath, dirfile1)
    if os.path.isfile(fileAllPath):
        file_name, file_ext = os.path.splitext(dirfile1)
        file_name_ALL, file_ext2 = os.path.splitext(fileAllPath)
        if file_ext == '.dwg' and file_name[-3:] != u'分层图' :
            fileNamea = file_name[:19]
            mkdir(file_name_ALL)  # 调用函数

filelista = []
for dirfile1 in dirs:
    fileAllPath = os.path.join(rootpath, dirfile1)
    if os.path.isfile(fileAllPath):
        file_name, file_ext = os.path.splitext(dirfile1)
        file_name_ALL, file_ext2 = os.path.splitext(fileAllPath)
        if file_ext == '.dwg' and file_name[-3:] != u'分层图':
            fileNamea = file_name[:19]
           # mkdir(file_name_ALL)  # 调用函数
            filelista.append(fileNamea)
            shutil.move(fileAllPath, file_name_ALL)
#        FileNewPath=os.path.join(file_name_ALL, file_name)
        elif file_ext == '.py':
            continue
        elif file_ext == '.dwg' and file_name[-3:] == u'分层图' :
            shutil.move(fileAllPath, file_name_ALL[:-3])
        elif file_ext == '.docx' :
            shutil.move(fileAllPath, file_name_ALL)
    else:
        continue
print len(list(set(filelista)))

