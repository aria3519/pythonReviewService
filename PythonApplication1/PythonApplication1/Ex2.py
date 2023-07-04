# 엑셀 데이터 구글스프레드 시트 저장
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
 
import gspread

from oauth2client.service_account import ServiceAccountCredentials
import selenium
from selenium.webdriver.support.select import Select
def WriteData():
    print()


# 한번에 주석 처리 컨트롤 kc // 해제 컨트롤 ku
def WebCrawlingSteam(LinkUrl,base):
      raw = requests.get(LinkUrl)
      html = BeautifulSoup(raw.text, 'html.parser')
      # summary -> Recent 변경
      # 옵션 변경 참고용
      #select=Select(driver.find_element_by_id("sch_bub_nm"))
      #select.select_by_index(1) #select index value
      #select.select_by_visible_text("Case2") # select visible text
      #select.select_by_value("000201") # Select option value
      
      select=Select(driver.find_element_by_id("review_context"))
      select.select_by_visible_text("recent") # select visible text

      container = html.select("div.rightcol")


      #for con in container:
      #    t = base #출처
      #    #print(con)
      #    c = con.select_one("div.forum_topic_name").text.strip() #글제목
      #    cou = con.select_one("div.forum_topic_reply_count").text.strip() #글코멘트수
      #    h = con.select_one("div.forum_topic_lastpost").text.strip() #글날짜
      #    l = 1
      #    worksheet.append_row([t, c, cou, h,l])# sheet 내 각 행에 데이터 추가
      #    print(t+"가 구글 스프레드에 최신화 되었습니다.")





# 구글 스프레드 시트 연동 
scope = [
    'https://spreadsheets.google.com/feeds'
    ,'https://www.googleapis.com/auth/drive'
]

# 가상 스마트 메일 json 파일 
json_file_name = "C://Users//appnori7//Desktop//google_sheet//smart-amplifier-390002-411995cffc5b.json"
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
gc = gspread.authorize(credentials)
print(gc)

# 연동 할려는 구글 스프레드 url
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1y0ipGFAf4j7ta-jHRzVMi5XYKQbaWjhrGZlhh9v894k/edit#gid=0'

# 스프레스시트 문서 가져오기
doc = gc.open_by_url(spreadsheet_url)
print(doc)

#시트 선택하기
worksheet = doc.worksheet('AIOReview')
print(worksheet)
#row_data = worksheet.row_values(1)
#print(row_data)
#range_list = worksheet.range('A1:M15')
#print(range_list)

worksheet.clear()
print("구글 스프레드 clear 되었습니다.")
# 4. 데이터프레임 내 header(변수명)생성
worksheet.append_row(["출처", "글제목","글코멘트수", "글날짜","글내용"])


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
    worksheet.append_row([t, c, cou, h,l])# sheet 내 각 행에 데이터 추가
print("구글 스프레드 최신화 되었습니다.")





WebCrawlingSteam("https://store.steampowered.com/app/1514840/AllInOne_Sports_VR/","steamReview")

