from win32com import client as wc
import os
from docx import Document
from docx.shared import Inches
import shutil
import re

words = "财务报表审计"

currentDir = os.getcwd()
# 原文件夹
originPath = os.path.join(currentDir, "files")
# 目标文件夹
targetName = 'fileTarget'
targetDir =  os.path.join(currentDir, targetName)

# 删除 fileTarget 文件夹
if os.path.exists(targetName):
    shutil.rmtree(targetName)

os.makedirs(targetName)

word = wc.Dispatch("Word.Application")

def transferFile(dir):
    filepath, filetype = os.path.splitext(dir)
    filename = os.path.basename(dir).replace(filetype, "")
    targetFile = targetDir + '\\' + filename + '.docx'
    # 文件名重复
    if os.path.isfile(targetFile):
        targetFile = os.path.join(targetDir, filename + '_1.docx')
    # doc类型转换
    if filetype == '.doc':
        doc = word.Documents.Open(dir)
        doc.SaveAs(targetFile, 12)
        doc.Close()
    # docx类型复制
    elif filetype == '.docx':
        shutil.copyfile(dir, targetFile)
    # 其他
    else:
        print("不支持处理的文件类型：" + dir)
    

# 遍历文件夹下所有文件
def eachFile(filepath):
    pathDir = os.listdir(filepath)
    for allDir in pathDir:
        child = os.path.join(filepath, allDir)
        # is directory
        if os.path.isdir(child):
            eachFile(child)
        else :
            transferFile(child)


eachFile(originPath)
word.Quit()




