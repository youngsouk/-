import pandas as pd
import docx 
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH ## 중앙정렬
from docx.enum.style import WD_STYLE_TYPE ## 스타일
### 단위 ###
from docx.shared import Mm
from docx.shared import Pt

import os, sys
#sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from poet_make import *

# run._element.rPr.rFonts.set(qn('w:eastAsia'), '휴먼명조') 한글 글꼴 설정

def rename(old_dict,old_name,new_name):
    new_dict = {}
    for key,value in zip(old_dict.keys(),old_dict.values()):
        new_key = key if key != old_name else new_name
        new_dict[new_key] = old_dict[key]
    return new_dict

p = pd.read_excel('./poet.xlsx')

poetries = p['시집']

titles_ex = p['제목']
contents = p['시']
titles = []

###title strip

for i in range(len(titles_ex)):
    titles.append(str(titles_ex[i]).strip())

### make title list to dict
titles_content = dict.fromkeys(titles, '')


for title,_ in titles_content.items():
    idx = titles.index(title)
    content = str(contents[idx])
    content = content.replace(title, '', 1)
    content = content.strip()

    if(titles_content[title] == ''):
        titles_content[title] += content
    if str(poetries[idx]) != 'nan':
        titles_content = rename(titles_content, title, title + ' [' + str(poetries[idx]) + ']')
'''   
### 내용 앞뒤 공백 제거
for i in range(len(contents)):
    contents[i] = str(contents[i])
    contents[i] = contents[i].replace(titles[i], '', 1)
    contents[i] = (contents[i]).strip()
'''
#doc = make_docx_list(title, content)

doc = make_docx_dict(titles_content)
doc.save('나희덕 시 전집.docx')
