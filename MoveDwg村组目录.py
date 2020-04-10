#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import shutil

# 读取中文路径 u''
#rootpath = unicode((os.getcwd()).replace("\\","/"),'gb2312')
rootpath=unicode(("D:\\1-p\\54、万店\\ALLinone").replace("\\","/"),'UTF-8')
dirs = os.listdir(rootpath)

def mkdir(path):
    folder = os.path.exists(path)
    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)  # makedirs 创建文件时如果路径不存在会创建这个路径
    else:
        print "---  There is this folder!  ---"

#创建文件夹
filelista = []
for dirfile1 in dirs:
    fileAllPath = os.path.join(rootpath, dirfile1)
    if os.path.isfile(fileAllPath):
        file_name, file_ext = os.path.splitext(dirfile1)
#        file_name_ALL, file_ext2 = os.path.splitext(fileAllPath)
        file_name_cun = file_name[6:12]
        file_name_zu = file_name[14:16]
        file_name_ALLpath = os.path.join(rootpath, file_name_cun, file_name_zu, file_name)
        if file_ext == '.dwg' and ( file_name[-3:] != u'分层图' and file_name[-5:] != u'分层平面图' ):
            mkdir(file_name_ALLpath)  # 调用函数
        elif file_ext == '.docx':
            mkdir(file_name_ALLpath)  # 调用函数
        else:
            continue

filelista = []
for dirfile1 in dirs:
    fileAllPath = os.path.join(rootpath, dirfile1)
    if os.path.isfile(fileAllPath):
        file_name, file_ext = os.path.splitext(dirfile1)
#        file_name_ALL, file_ext2 = os.path.splitext(fileAllPath)
        file_name_cun = file_name[6:12]
        file_name_zu = file_name[14:16]
        file_name_ALLpath = os.path.join(rootpath, file_name_cun, file_name_zu, file_name)
        if file_ext == u'.dwg' and (u'分层图' == file_name[-3:]  ) :
            shutil.move(fileAllPath, os.path.join(file_name_ALLpath[:-3], dirfile1))
        elif file_ext == u'.docx':
            shutil.move(fileAllPath, file_name_ALLpath)
        elif file_ext == '.dwg' and file_name[-3:] != u'分层图':
            filelista.append(file_name)
            shutil.move(fileAllPath, file_name_ALLpath)
        else:
            continue
    else:
        continue

print len(list(set(filelista)))

