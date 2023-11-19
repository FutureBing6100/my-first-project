#DocxTemplate批量生成word.py
import os

import pandas as pd
from docxtpl import DocxTemplate

#一、导入相关模块，设定excel所在文件夹和生成word保存的文件夹
zpath=os.getcwd()+'\\'
zpath=r'C:\Users\phoen\Desktop\pp'+'\\'
#系统路径如下面的路径，使用r就防止了\t的转义
file_path=zpath+r'\通知单合集'

#二、遍历excel，逐个生成word（form.docx是前面的模板）
try:
    os.mkdir(file_path)
except:
    pass
#os.mkdir创建目录，如果已经有了，跳过


tpl = DocxTemplate(zpath+'form.docx')
autho = pd.read_excel(zpath+'autho.xlsx')
name = autho["name"].str.rstrip()
classs = autho['classs'].str.rstrip()  # str.rstrip()用于去掉换行符
chi = autho['chi']
math = autho['math']
eng = autho['eng']

# 遍历excel行，逐个生成
num = autho.shape[0]
for i in range(num):
    context = {
       "name": name[i],
       "classs": classs[i],
       "chi": chi[i],
       "math": math[i],
       "eng": eng[i]
    }
    tpl = DocxTemplate(zpath+'form.docx')
    tpl.render(context)
    tpl.save(file_path+r"\{}的成绩通知单.docx".format(name[i]))
