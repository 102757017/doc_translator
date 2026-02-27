# -*- coding: UTF-8 -*-
import os
import sys
import shutil
import re
import xlwings as xls
import pprint
import translators as ts


if getattr(sys, 'frozen', False):
    bundle_dir = sys._MEIPASS
else:
    bundle_dir = os.path.dirname(os.path.abspath(__file__))

os.chdir(bundle_dir)

print(ts.translators_pool)

app = xls.App(visible=True,add_book=False)
#要翻译的文档
file_path="QAV-1.xlsx"
workbook=app.books.open(file_path)
for sht in app.books[file_path].sheets:
    max_row=sht.used_range.rows.count
    max_column=sht.used_range.columns.count
    for row in range(max_row):
        for column in range(max_column):
            orgin=sht[row,column].value
            if isinstance(orgin, str) and len(orgin)>=2:
                #只翻译英文单词
                flag = re.fullmatch(r"[A-Za-z][A-Za-z]*",orgin)
                if flag!=0:
                    print("原文",orgin)
                    result = ts.translate_text(orgin, to_language='zh')
                    print("译文",result)
                    sht[row,column].value=result
print(">>>>>>>>>>----------翻译结束----------<<<<<<<<<<")                
            
#写入excel
workbook.save('result.xlsx')
workbook.close()

    
