import requests
from bs4 import BeautifulSoup
import openpyxl


from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
 

# 2. Workbook 생성
wb = openpyxl.Workbook()

# 3. Sheet 활성
sheet = wb.active

# 4. 데이터프레임 내 header(변수명)생성
sheet.append(["출처", "글제목","글코멘트수", "글날짜","글내용"])

# 5. 데이터 크롤링 과정
raw = requests.get("https://steamcommunity.com/app/1514840/discussions/")
html = BeautifulSoup(raw.text, 'html.parser')

container1 = html.select("a.forum_topic_overlay")
#container1 = html.select("a.href")
#container = html.select("div.forum_topics_container")
container = html.select("div.forum_topic")

#print(container1)

listcontent = list()
# 링크 가져 오고 그 안에 내용 가져오면 됨
for con in container1:
    content = con.attrs['href'] # 링크 가져옴
    #print(content)
    contentUrl = requests.get(content)
    contentHtml = BeautifulSoup(contentUrl.text, 'html.parser')
    contentCont = contentHtml.select_one("div.forum_op")
    contentContCont = contentCont.select_one("div.content").text.strip()
    #print(contentContCont)
    listcontent.append(contentContCont)
    
i=0
for con in container:
    t = "steamcommunity" #출처
    #print(con)
    c = con.select_one("div.forum_topic_name").text.strip() #글제목
    cou = con.select_one("div.forum_topic_reply_count").text.strip() #글코멘트수
    h = con.select_one("div.forum_topic_lastpost").text.strip() #글날짜
    #l =con.select_one("span.topic_hover_data") #글내용
    #print(l)
    l = listcontent[i]
    #print(l)
    i+=1
    #print(i)
    sheet.append([t, c, cou, h,l])# sheet 내 각 행에 데이터 추가



# 각 칼럼에 대해서 모든 셀값의 문자열 개수에서 1.1만큼 곱한 것들 중 최대값을 계산한다.
for column_cells in sheet.columns:
    length = max(len(str(cell.value))*1.1 for cell in column_cells)
    sheet.column_dimensions[column_cells[0].column_letter].width = length
    ## 셀 가운데 정렬
    for cell in sheet[column_cells[0].column_letter]:
        cell.alignment = Alignment(horizontal='center')

# 6. 수집한 데이터 저장
wb.save(r"C://Users\appnori7//Desktop//sa//CheckReview.xlsx")
wb.close()
print("최신화 되었습니다.")
