import requests
from bs4 import BeautifulSoup
import openpyxl

# 2. Workbook 생성
wb = openpyxl.Workbook()

# 3. Sheet 활성
sheet = wb.active

# 4. 데이터프레임 내 header(변수명)생성
sheet.append(["제목", "채널명", "조회수", "좋아요수"])

# 5. 데이터 크롤링 과정
raw = requests.get("https://tv.naver.com/r")
html = BeautifulSoup(raw.text, 'html.parser')

container = html.select("div.inner")
#print(container)
for con in container:
        t = con.select_one("dt.title").text.strip() #제목
        c = con.select_one("dd.chn").text.strip() #채널
        h = con.select_one("span.hit").text.strip() #조회수
        l = con.select_one("span.like").text.strip() #좋아요수
    
 # sheet 내 각 행에 데이터 추가
        sheet.append([t, c, h, l])


# 6. 수집한 데이터 저장
wb.save(r"C://Users\appnori7//Desktop//sa//naver_tv.xlsx")
print("최신화 되었습니다.")
