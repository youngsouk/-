import pandas as pd
import docx 
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH ## 중앙정렬
from docx.enum.style import WD_STYLE_TYPE ## 스타일
### 단위 ###
from docx.shared import Mm
from docx.shared import Pt

# run._element.rPr.rFonts.set(qn('w:eastAsia'), '휴먼명조') 한글 글꼴 설정

p = pd.read_excel('./poet.xlsx')

titles = p['title']
contents = p['poet']

###title strip
print (len(titles))
for i in range(len(titles)):
    titles[i] = str(titles[i]).strip()
    
### 내용 앞뒤 공백 제거
for i in range(len(contents)):
    contents[i] = str(contents[i])
    contents[i] = contents[i].replace(titles[i], '')
    contents[i] = (contents[i]).strip()

doc = docx.Document()
doc.add_heading('나희덕 시 전집', 0)

### 문서 기본 설정 ###
section = doc.sections[0]
section.page_width = Mm(148)
section.page_heiht = Mm(225)

section.top_margin = Mm(10)
section.bottom_margin = Mm(10)

section.left_margin = Mm(15)
section.right_margin = Mm(15)

section.header_distance = Mm(10)
section.footer_distance = Mm(10)
##################

### 문서 스타일 생성 ###
styles = doc.styles

title_style = styles.add_style('시 제목',WD_STYLE_TYPE.PARAGRAPH)
title_style.font.name = '나눔고딕 ExtraBold'
title_style._element.rPr.rFonts.set(qn('w:eastAsia'), '나눔고딕 ExtraBold')
title_style.font.size = Pt(12)

content_style = styles.add_style('시 본문',WD_STYLE_TYPE.PARAGRAPH)
content_style.font.name = '나눔 명조'
content_style._element.rPr.rFonts.set(qn('w:eastAsia'), '나눔 명조')
content_style.font.size = Pt(11)
####################

for i in range(len(titles)):
    ### 제목 ###
    doc.add_paragraph(titles[i] + '\n', style = title_style).bold = True
    last_paragraph = doc.paragraphs[-1] 
    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    #### 내용 ###
    doc.add_paragraph(contents[i] + '\n\n', style = content_style)
    
    ### 페이지 넘어가기 ####
    doc.add_page_break()

doc.save('나희덕 시 전집.docx')

#### 제목.txt 생성
'''
idx = 0
for title in titles:
    if '/' not in str(title):
        f = open('./poets/'+str(title) + '.txt', mode = 'w')
        f.write(str(contents[idx]))
        idx +=1
    else:
        print (title)
        idx += 1
'''