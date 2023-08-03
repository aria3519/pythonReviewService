# 엑셀 데이터 구글스프레드 시트 저장
import requests
from bs4 import BeautifulSoup
# pip install BeautifulSoup4
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
 
import gspread

from oauth2client.service_account import ServiceAccountCredentials


import selenium
import chromedriver_autoinstaller
from selenium import webdriver 
#from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.select import Select

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

import numpy as np


from flask import Flask,request,jsonify
import json
import PyKakao
from PyKakao import Message
from datetime import datetime

from telegram import Bot
import sys
import asyncio


#import time, win32con, win32api, win32gui
tel_token = 'bot6150221677:AAE4Ue-beARKbFtM1472toqPMf-QCtjpr5E'
tel_id = '-1001838059229' # 내 아이디
def tel_id_check():
    chat = telegram.Bot(token = tel_token)
    updates = chat.getUpdates()
    for i in updates:
        print(u.message['chat']['id'])

def SendTelegramBot(title="  ",content ="",day = "",point="",gameType=""):

    data = {"chat_id" : tel_id, "text": f"게임: "+gameType+f"\n날짜: "+day + f"\n출처: "+title+"평점:"+point +"\n내용:"+content }
    res = requests.post('https://api.telegram.org/bot6150221677:AAE4Ue-beARKbFtM1472toqPMf-QCtjpr5E/sendMessage',
                        data = data)
    print(res)

def SendAlarm(alarmList,index,data,notCheck,title,day,point,gametype=0,eng=""):
    if(notCheck == True or index >= len(alarmList)):
        return "Old"
    if(alarmList[index] == "-1"):
        return "Old"
    if(alarmList[index] != data ):
        #SendMsgForKakao(title,data)
        #SendTelegramBot(title,data)
        
        ty =" "
        if(gametype == 0):
            ty = "All_In_One"
        elif(gametype ==1):
             ty = "All_In_One_Summer"
        elif(gametype ==2):
            ty = "RealGostopVR"
        if(eng ==""):
            SendTelegramBot(title,data,day,point,ty)
        else:
            SendTelegramBot(title,eng,day,point,ty)
        
        
        ##print(title+"\w"+day+"\w"+point+"\w"+eng)
        return "New"
    else:
        return "Old"
def FindButten(driver,type, value):
    but = WebDriverWait(driver,10).until(EC.element_to_be_clickable((type,value)))
    but.click()
    
def FindButten2(driver,type, value):
    but = WebDriverWait(driver,10).until(EC.element_to_be_clickable((type,value)))
    driver.execute_script("arguments[0].click()",but)





# 한달에 한번 refreshtoken 생성
def CreateKakaoJson():

    api = Message(service_key = "60ebe15145bad656cff8f17a71e888af")
    auth_url = api.get_url_for_generating_code()
    
    #access_token = api.get_access_token_by_redirected_url(auth_url)
    #인증코드 받는 주소 
    #https://kauth.kakao.com/oauth/authorize?client_id=60ebe15145bad656cff8f17a71e888af&redirect_uri=https://appnoriReview.com/oauth&response_type=code&scope=talk_message,friends

    kakaourl = "https://kauth.kakao.com/oauth/token"
    client_id = '60ebe15145bad656cff8f17a71e888af'
    redirect_uri = 'https://appnoriReview.com/oauth'
    code = GetKakaoCode()
    #code = '1JQZuxNu53DWBz4xjA3NNSHc0BWUFoCGhCtAFywlWk-Fq3Izbfww4t-N3fiYeqcoUeHkPAo9dJgAAAGJHrKXfg'
    
    data = {
        'grant_type':'authorization_code',
        'client_id':client_id,
        'redirect_uri':redirect_uri,
        'code': code,
        }
    response = requests.post(kakaourl, data=data)
    tokens = response.json()
    with open("token.json","w") as kakao:
        json.dump(tokens, kakao)
    print("초기 인증 토큰 저장 성공")
    
    
    with open("token.json", "r") as fp:
        tokens = json.load(fp)
        
    data = {
        "grant_type": "refresh_token",
        "client_id": client_id,
        "refresh_token": tokens['refresh_token']
        }
    response = requests.post(kakaourl, data=data)
    result = response.json()
    print(tokens)
    if 'access_token' in result:
        tokens['access_token'] = result['access_token']
    if 'refresh_token' in result:
        tokens['refresh_token'] = result['refresh_token']
    else:
        pass

    #발행된 토큰 저장
    with open("token.json","w") as kakao:
        json.dump(tokens, kakao)
    print("발행된 refresh_token 저장")

# 인증 기간 긴 토큰 생성 
def CreateKakaoRefresh(r_token):
    with open("token.json", "r") as fp:
        ts = json.load(fp)
    data = {
        "grant_type": "refresh_token",
        "client_id":'60ebe15145bad656cff8f17a71e888af',
        "refresh_token": r_token
    }
    kakaourl = "https://kauth.kakao.com/oauth/token"
    response = requests.post(kakaourl, data=data)
    tokens = response.json()

    with open(r"token.json", "w") as fp:
        json.dump(tokens, fp)
    with open("token.json", "r") as fp:
        ts = json.load(fp)
    token = ts["access_token"]
    return token


def SendMsgForKakao(title="  ",content =""):
    with open("token.json","r") as fp:
        ts = json.load(fp)
    tokens = ts["refresh_token"]
    data = {
        "grant_type": "refresh_token",
        "client_id":'60ebe15145bad656cff8f17a71e888af',
        "refresh_token": tokens
    }
    kakaourl = "https://kauth.kakao.com/oauth/token"
    response = requests.post(kakaourl, data=data)
    tokens = response.json()
    with open("token.json", "r") as fp:
        ts = json.load(fp)
    token = ts["access_token"]
    #friend_url = "https://kapi.kakao.com/v2/user/me"
    #headers={"Authorization" : "Bearer " + "eo2ttJDzF2pUXyEWudbnwcuSmu3KoJqL_gtMkSMwCiolkAAAAYkZfJh5"}
    friend_url = "https://kapi.kakao.com/v1/api/talk/friends"
    headers={"Authorization" : "Bearer " +  token}
    result = json.loads(requests.get(friend_url, headers=headers).text)
    print(type(result))
    print("=============================================")
    print(result)
    print("=============================================")
    friends_list = result.get("elements")
    print(friends_list)
    # print(type(friends_list))
    print("=============================================")
    #print(friends_list[0].get("uuid"))
    friend_id = friends_list[1].get("uuid")
    friend_id2 = friends_list[0].get("uuid")
    print(friend_id)
  
    send_url= "https://kapi.kakao.com/v1/api/talk/friends/message/default/send"
    #{'elements': [{'profile_nickname': '김민석', 'profile_thumbnail_image': '', 'allowed_msg': True, 'id': 2884442397, 'uuid': 'SXpLfEh_THlBbVttVWZebV1oRH1OfE58RC0', 'favorite': False}], 'total_count': 1, 'after_url': None, 'favorite_count': 0}
    # {'profile_nickname': '이 상욱 부대표님 앱노리', 'profile_thumbnail_image': 'https://p.kakaocdn.net/th/talkp/wlkJcQ87Jq/mJcIzQRBKXp1kWbNkfAEK1/st1zdp_110x110_c.jpg', 'allowed_msg': True, 'id': 2888646220, 'uuid': 'SXtLfU5_T3lBbVxvW29faV1xSHtJe0lxHg', 'favorite': False}
    # {'profile_nickname': '이현욱 대표님 앱노리', 'profile_thumbnail_image': 'https://p.kakaocdn.net/th/talkp/wns8qTVlCj/9MeJSJxY1rV8r9Vxb2fTBk/h7wt6q_110x110_c.jpg', 'allowed_msg': True, 'id': 2885848675, 'uuid': 'SXpJekl5TXxIZFdkUmVTYFluQntIekh6Qh4', 'favorite': False}], 'total_count': 2, 'after_url': None, 'favorite_count': 0}
    #uuidsData = {"receiver_uuids" : json.dumps(friend_id)}    
    #-data-urlencode 'receiver_uuids=["abcdefg0001","abcdefg0002","abcdefg0003"]'
    data={
    'receiver_uuids':'["{}","{}"]'.format(friend_id,friend_id2),
    "template_object": json.dumps({
        "object_type":"text",
        "text":f"출처: "+title+"\n내용:"+content,
        "link":{
            "web_url" :"https://docs.google.com/spreadsheets/d/1y0ipGFAf4j7ta-jHRzVMi5XYKQbaWjhrGZlhh9v894k/edit#gid=0",
            "mobile_web_url" :"https://docs.google.com/spreadsheets/d/1y0ipGFAf4j7ta-jHRzVMi5XYKQbaWjhrGZlhh9v894k/edit#gid=0",
        },
        "button_title": "확인"
        })
    }
   
    response = requests.post(send_url, headers=headers, data=data)
    print(response)
    print(response.json())
    if(response.status_code != 200 and content ==""):
        print("카카오톡 api 에러")
        CreateKakaoJson()
        #sys.exit()
   

def GetKakaoCode():
    driver = webdriver.Chrome()
    driver.get("https://kauth.kakao.com/oauth/authorize?client_id=60ebe15145bad656cff8f17a71e888af&redirect_uri=https://appnoriReview.com/oauth&response_type=code&scope=talk_message,friends")
    time.sleep(5)
    inputid = driver.find_element(By.XPATH,value ='//*[@id="loginKey--1"]')
    inputpw = driver.find_element(By.XPATH,value ='//*[@id="password--2"]')
    inputid.send_keys("aria3519@naver.com")
    inputpw.send_keys("destiny3519!!")
    time.sleep(2)
    but = driver.find_element(By.XPATH,value ='//*[@id="mainContent"]/div/div/form/div[4]/button[1]')
    but.click()
    time.sleep(10)
    url=driver.current_url
    print(driver.current_url)
    url = url.replace("https://appnorireview.com/oauth?code=","")
    print(url)
    return url


def WebCrawlingSteamReviewCommu():
    raw = requests.get("https://steamcommunity.com/app/2074750/discussions/")
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
        contenturl = requests.get(content)
        contenthtml = BeautifulSoup(contenturl.text, 'html.parser')
        contentcont = contenthtml.select_one("div.forum_op")
        contentcontcont = contentcont.select_one("div.content").text.strip()
        #print(contentcontcont)
        listcontent.append(contentcontcont)
    


    i=0
    index=4
    new = "Old"
    for con in container:
    
        t = "SteamCommunity" #출처
        total = "x"
        point = "x"
        c = con.select_one("div.forum_topic_name").text.strip() #글제목
        cou = con.select_one("div.forum_topic_reply_count").text.strip() #글코멘트수
        h = con.select_one("div.forum_topic_lastpost").text.strip() #글날짜
        l = listcontent[i]
        check = False
        if(index>=5):
            check = True
        if(new == "New"):
            new = SendAlarm(alarmList,index,l,check,t,h,point,1)
        else:
            new = SendAlarm(alarmList,index,l,check,t,h,point,1)
        i+=1
        if(new !="New"):
            index+=1
        worksheet.append_row([new,t, c,total,point, cou, h,l])# sheet 내 각 행에 데이터 추가

# 한번에 주석 처리 컨트롤 kc // 해제 컨트롤 ku
def WebCrawlingSteamReview(url,base,worksheet):
      #raw = requests.get(url)
      #html = BeautifulSoup(raw.text, 'html.parser')
      # 좋아요 싫어요 평점에 넣고 전체 평점 받아오기
      # 글 제목에는 아이디 넣기 
    chrome_ver = chromedriver_autoinstaller.get_chrome_version()
    #print("chrome_ver: "+chrome_ver)
    path=chromedriver_autoinstaller.install()
    driver = webdriver.Chrome()
    #print("driver_ver: "+driver.capabilities['browserVersion'])
    driver.get(url)
    # summary -> Recent 변경
    select=Select(driver.find_element(By.CSS_SELECTOR,value ="#review_context"))
    #select.select_by_index(1) #select index value
    #"user_reviews_filter_display_as"
    select.select_by_index(2)
    time.sleep(10)
      #container = driver.find_elements(By.CLASS_NAME,value = "rightcol")
      #print(container)
    raw = driver.page_source
    html = BeautifulSoup(raw, 'html.parser') 
    container = html.select('div#Reviews_recent>div>div.review_box  ')
      #print(container)
      #option = html.select("div.review_developer_response_container.multiple_listing.store")
      #option=html.select("a.vote_header.tooltip")
      #print(option)
    total =driver.find_element(By.XPATH,value ='//*[@id="review_histogram_rollup_section"]/div[1]/div/span[1]').text.strip()
    i = 2
    if(url == "https://store.steampowered.com/app/2074750/AllInOne_Summer_Sports_VR/"):
          i=5
    new = "Old"
    for con in container:
        t = base #출처
        point = con.select_one("div.title.ellipsis").text.strip() #좋아요/싫어요
        c = con.select_one("div.persona_name").text.strip() #글제목
        cou = con.select_one("div.vote_info").text.strip() #글코멘트수
        h = con.select_one("div.postedDate").text.strip() #글날짜
        l = con.select_one("div.content").text.strip() #글내용
        check = False
        if(i>=4 and url == "https://store.steampowered.com/app/1514840/allinone_sports_vr/"):
              check = True
        if(i>=6 and url == "https://store.steampowered.com/app/2074750/AllInOne_Summer_Sports_VR/"):
              check = True
        if(new == "New"):
            if(url == "https://store.steampowered.com/app/2074750/AllInOne_Summer_Sports_VR/"):
                new = SendAlarm(alarmList,i,l,check,base,h,point,1)
            else:
                new = SendAlarm(alarmList,i,l,check,base,h,point)
        else:
            if(url == "https://store.steampowered.com/app/2074750/AllInOne_Summer_Sports_VR/"):
                     new = SendAlarm(alarmList,i,l,check,base,h,point,1)
            else:
                new = SendAlarm(alarmList,i,l,check,base,h,point)
        if(new !="New"):
            i+=1
        worksheet.append_row([new,t, c,total,point, cou, h,l])# sheet 내 각 행에 데이터 추가
    driver.close()
    
      #for con in container:
      #    t = base #출처
      #    #print(con)
      #    c = con.select_one("div.forum_topic_name").text.strip() #글제목
      #    cou = con.select_one("div.forum_topic_reply_count").text.strip() #글코멘트수
      #    h = con.select_one("div.forum_topic_lastpost").text.strip() #글날짜
      #    l = 1
      #    worksheet.append_row([t, c, cou, h,l])# sheet 내 각 행에 데이터 추가
      #    print(t+"가 구글 스프레드에 최신화 되었습니다.")


def WebCrawlingPico(urlLog,url1,url2,base,worksheet):
    chrome_ver = chromedriver_autoinstaller.get_chrome_version()
    path=chromedriver_autoinstaller.install()
    driver = webdriver.Chrome()
    driver.get(urlLog)
    time.sleep(10)
    # 로그인하기
    #select=Select(driver.find_element(By.CSS_SELECTOR,value ="#review_context"))

    inputid = driver.find_element(By.XPATH,value ='//*[@id="root"]/main/div/div/div/article/article/article/form/div[1]/div[2]/input')
    inputpw = driver.find_element(By.XPATH,value ='//*[@id="root"]/main/div/div/div/article/article/article/form/div[2]/div/input')
    inputid.send_keys("howard@appnori.com")
    inputpw.send_keys("Appnori73")
    time.sleep(10)
    FindButten(driver,By.XPATH,'//*[@id="root"]/main/div/div/div/article/article/article/form/div[5]/button')
    #but = driver.find_element(By.XPATH,value ='//*[@id="root"]/main/div/div/div/article/article/article/form/div[5]/button')
    #but.click()
    time.sleep(10)
    FindButten(driver,By.XPATH,'//*[@id="dev-warp"]/div/div[3]/div[3]/div/div[1]/button')
    #butx = driver.find_element(By.XPATH,value ='//*[@id="dev-warp"]/div/div[3]/div[3]/div/div[1]/button')
    #butx.click()
    time.sleep(10)
    # pico 리뷰 페이지 
    #driver.get("https://developer-global.pico-interactive.com/console#/app/reviews/397/7098225807675359237")
    driver.get(url1)
    time.sleep(10)
    FindButten(driver,By.XPATH,'//*[@id="pane-1"]/div/div[2]/div[1]/div[1]/div[1]/div/span/span')
    #butRecent = driver.find_element(By.XPATH,value ='//*[@id="pane-1"]/div/div[2]/div[1]/div[1]/div[1]/div/span/span')
    #butRecent.click()
    FindButten2(driver,By.XPATH,'/html/body/div[3]/div[1]/div[1]/ul/li[3]/span')
    #butRecenttime = driver.find_element(By.XPATH,value ='/html/body/div[3]/div[1]/div[1]/ul/li[3]/span')
    #driver.execute_script("arguments[0].click()",butRecenttime)
    
    #butRecenttime = driver.find_element(By.XPATH,value ='/html/body/div[3]/div[1]/div[1]/ul')
    #driver.execute_script("arguments[0].click()",butRecenttime)
    #/html/body/div[5]/div[1]/div[1]/ul/li[2]
    time.sleep(10)
    raw = driver.page_source
    html = BeautifulSoup(raw, 'html.parser')
    time.sleep(10)
    container = html.select('div.review_card')
    checklist = list()
    total =  html.select_one('span.number').text.strip()
    for i in range(0,20):
         checklist.append(container[i])
    i = 4
    new = "Old"
    for con in checklist:
        t = base+"GB" #출처
        c = con.select_one("div.header>div>div>span.name").text.strip()
        pointlist = con.select('img')
        count = 0
        for point in pointlist:
            if(point.get('src') =="https://sf16-scmcdn-va.ibytedtos.com/obj/static-us/pico/developer_frontend/img/rating_star_yellow.0a718ebc.svg"):
                count += 1
        cou = "X"#글코멘트수
        h = con.select_one("div.header>div>div>span.time").text.strip() #글날짜
        l = con.select_one("div.content>div.review").text.strip() #글내용
        check = False
        if(i>=6):
            check = True
        if(new == "New"):
            new = SendAlarm(alarmList,i,l,check,t,h,str(count))
        else:
            new = SendAlarm(alarmList,i,l,check,t,h,str(count))
        if(new !="New"):
            i+=1
        worksheet.append_row([new,t,c,total,str(count), cou, h,l])# sheet 내 각 행에 데이터 추가
        time.sleep(1)
         
             
    # 중국 리뷰 
    driver.get(url2)
    time.sleep(10)
    FindButten(driver,By.XPATH,'//*[@id="tab-0"]')
    #but = driver.find_element(By.XPATH,value ='//*[@id="tab-0"]')
    #but.click()
    time.sleep(10)
    FindButten(driver,By.XPATH,'//*[@id="pane-0"]/div/div[2]/div[1]/div[1]/div[1]/div/span/span/i')
    #butRecent = driver.find_element(By.XPATH,value ='//*[@id="pane-0"]/div/div[2]/div[1]/div[1]/div[1]/div/span/span/i')
    #butRecent.click()
    time.sleep(10)
    FindButten2(driver,By.XPATH,'/html/body/div[3]/div[1]/div[1]/ul/li[2]')
    #butRecenttime = driver.find_element(By.XPATH,value ='/html/body/div[3]/div[1]/div[1]/ul/li[2]')
    #driver.execute_script("arguments[0].click()",butRecenttime)
    time.sleep(10)
    raw = driver.page_source
    html = BeautifulSoup(raw, 'html.parser')
    containerChina = html.select('div.review_card')
    FindButten(driver,By.XPATH,'//*[@id="pane-0"]/div/div[2]/div[1]/div[2]/div/div/span/span')
    #butShow = driver.find_element(By.XPATH,value ='//*[@id="pane-0"]/div/div[2]/div[1]/div[2]/div/div/span/span')
    #butShow.click()
    time.sleep(10)
    FindButten2(driver,By.XPATH,'/html/body/div[4]/div[1]/div[1]/ul/li[2]/span')
    #butEng = driver.find_element(By.XPATH,value ='/html/body/div[4]/div[1]/div[1]/ul/li[2]/span')
    #driver.execute_script("arguments[0].click()",butEng)
    time.sleep(10)
    raw = driver.page_source
    html = BeautifulSoup(raw, 'html.parser')
    containerEng = html.select('div.review_card')
    checklist.clear()
    for i in range(0,20):
         checklist.append(containerChina[i])
    index = 0
    i = 6
    total =  html.select_one('span.number').text.strip()
    for con in checklist:
        t = base+"China" #출처
        new ="X"
          #print(con)
        c = con.select_one("div.header>div>div>span.name").text.strip() #글제목
        pointlist = con.select('img')
        count = 0
        for point in pointlist:
            if(point.get('src') =="https://sf16-scmcdn-va.ibytedtos.com/obj/static-us/pico/developer_frontend/img/rating_star_yellow.0a718ebc.svg"):
                count += 1
        cou = "X"#글코멘트수
        h = con.select_one("div.header>div>div>span.time").text.strip() #글날짜
        l = con.select_one("div.content>div.review").text.strip() #글내용
        eng = containerEng[index].select_one("div.content>div.review").text.strip() # 번역 내용
        check = False
        if(i>=8):
            check = True
        if(new == "New"):
            new = SendAlarm(alarmList,i,l,check, t,h,str(count),0,eng)
        else:
            new = SendAlarm(alarmList,i,l,check, t,h,str(count),0,eng)
        if(new !="New"):
            i+=1
        index += 1
        worksheet.append_row([new,t,c,total,str(count), cou, h,l,eng])# sheet 내 각 행에 데이터 추가
        time.sleep(1)
    
    driver.close()




def WebCrawlingOculus1(url,base,worksheet):
    chrome_ver = chromedriver_autoinstaller.get_chrome_version()
    path=chromedriver_autoinstaller.install()
    driver = webdriver.Chrome()
    driver.get(url)
    time.sleep(10)
    #//*[@id="addReviewBox"]/div/div[3]/div/div/div/div[1]/div[2]/div[2]
    FindButten2(driver,By.XPATH,'//*[@id="addReviewBox"]/div/div[3]/div/div/div/div[1]/div[2]/div[2]')
    #butEng = driver.find_element(By.XPATH,value ='//*[@id="addReviewBox"]/div/div[3]/div/div/div/div[1]/div[2]/div[2]')
    #driver.execute_script("arguments[0].click()",butEng)
    time.sleep(10)
    raw = driver.page_source
    html = BeautifulSoup(raw, 'html.parser')
    container = html.select('div.rpc-content.ng-star-inserted')
    total = html.select_one('div.counter.ng-star-inserted').text.strip()
    i = 9
    for con in container:
        t = base #출처
        new ="X"
        try:
            c = con.select_one("div.user-name.condensed.cursor-pointer.truncate").text.strip() #글제목
        except Exception as e:
            continue
          #pointtemp = con.select_one('div.stars-wrapper.right')#글날짜
        pointtemp = con.select_one('sq-ratings-stars')#글날짜
        point = pointtemp.get('rating')
        cou = "X"#글코멘트수
        h = con.select_one("div.rpc-date.inline-block").text.strip() #글날짜
        l = con.select_one("div.small-padding>p").text.strip() #글내용
        check = False
        if(i>=10):
            check = True
        new = SendAlarm(alarmList,i,l,check,base,h,point)
        i += 1
        worksheet.append_row([new,t, c,total,point, cou, h,l])# sheet 내 각 행에 데이터 추가
        time.sleep(1)
    
    

    driver.close()






def WebCrawlingOculus2(url,base,worksheet):
    chrome_ver = chromedriver_autoinstaller.get_chrome_version()
    path=chromedriver_autoinstaller.install()
    driver = webdriver.Chrome()
    driver.get(url)
    #driver.back()
    #driver.forward()
    #driver.refresh()
    time.sleep(10)
    FindButten(driver,By.XPATH,'//*[@id="facebook"]/body/div[2]/div/div[2]/div[1]/div[2]/i')
    #butx = driver.find_element(By.XPATH,value ='//*[@id="facebook"]/body/div[2]/div/div[2]/div[1]/div[2]/i')
    #butx.click()
    time.sleep(10)
    FindButten2(driver,By.CSS_SELECTOR,"#mount > div > main > div > div > div > div.app__content > div.app__info > div > div.app__description > div.app__reviews > div > div.app-review-list > div.app-review-list__sort-filters > span:nth-child(1) > a")
    #butShow = driver.find_element(By.CSS_SELECTOR,value="#mount > div > main > div > div > div > div.app__content > div.app__info > div > div.app__description > div.app__reviews > div > div.app-review-list > div.app-review-list__sort-filters > span:nth-child(1) > a")
    #driver.execute_script("arguments[0].click()",butShow)
    time.sleep(10)
    #body > div:nth-child(28) > div > ul > li:nth-child(2)
    butRecent = driver.find_elements(By.CLASS_NAME,value ="sky-dropdown__item")
    for but in butRecent:
        if(but.text=="정렬: 최신순"):
            but.click()
    time.sleep(10)
    raw = driver.page_source
    html = BeautifulSoup(raw, 'html.parser')
    container = html.select('div.app-review')
    total = "X"
    i = 8 
    for con in container:
        t = base #출처
        c = con.select_one("h1.bxHeading.bxHeading--level-5.app-review__title").text.strip() #글제목
        pointlist = con.select("i.bxStars.bxStars--white")
        point = str(len(pointlist))
        #cou = con.select_one("div.footer>div.likenum>div.like>span")#글코멘트수
        cou = "X"#글코멘트수
        h = con.select_one("div.app-review__date").text.strip() #글날짜
        l = con.select_one("div.clamped-description__content").text.strip() #글내용
        check = False
        if(i>=9):
            check = True
        new = SendAlarm(alarmList,i,l,check,base,h,point)
        i += 1
        worksheet.append_row([new,t, c,total,point, cou, h,l])# sheet 내 각 행에 데이터 추가
        time.sleep(1)

def WebCrawlingPicoSummer(urlLog,url1,url2,base):
    chrome_ver = chromedriver_autoinstaller.get_chrome_version()
    path=chromedriver_autoinstaller.install()
    driver = webdriver.Chrome()
    driver.get(urlLog)
    time.sleep(10)
    # 로그인하기
    #select=Select(driver.find_element(By.CSS_SELECTOR,value ="#review_context"))

    inputid = driver.find_element(By.XPATH,value ='//*[@id="root"]/main/div/div/div/article/article/article/form/div[1]/div[2]/input')
    inputpw = driver.find_element(By.XPATH,value ='//*[@id="root"]/main/div/div/div/article/article/article/form/div[2]/div/input')
    inputid.send_keys("howard@appnori.com")
    inputpw.send_keys("Appnori73")
    time.sleep(10)
    FindButten(driver,By.XPATH,'//*[@id="root"]/main/div/div/div/article/article/article/form/div[5]/button')
    #but = driver.find_element(By.XPATH,value ='//*[@id="root"]/main/div/div/div/article/article/article/form/div[5]/button')
    #but.click()
    time.sleep(10)
    FindButten(driver,By.XPATH,'//*[@id="dev-warp"]/div/div[3]/div[3]/div/div[1]/button')
    #butx = driver.find_element(By.XPATH,value ='//*[@id="dev-warp"]/div/div[3]/div[3]/div/div[1]/button')
    #butx.click()
    time.sleep(10)
    # pico 리뷰 페이지 
    #driver.get("https://developer-global.pico-interactive.com/console#/app/reviews/397/7098225807675359237")
    driver.get(url1)
    time.sleep(10)
    FindButten(driver,By.XPATH,'//*[@id="pane-1"]/div/div[2]/div[1]/div[1]/div[1]/div/span/span')
    #butRecent = driver.find_element(By.XPATH,value ='//*[@id="pane-1"]/div/div[2]/div[1]/div[1]/div[1]/div/span/span')
    #butRecent.click()
    FindButten2(driver,By.XPATH,'/html/body/div[3]/div[1]/div[1]/ul/li[3]/span')
    #butRecenttime = driver.find_element(By.XPATH,value ='/html/body/div[3]/div[1]/div[1]/ul/li[3]/span')
    #driver.execute_script("arguments[0].click()",butRecenttime)
    
    #butRecenttime = driver.find_element(By.XPATH,value ='/html/body/div[3]/div[1]/div[1]/ul')
    #driver.execute_script("arguments[0].click()",butRecenttime)
    #/html/body/div[5]/div[1]/div[1]/ul/li[2]
    time.sleep(10)
    raw = driver.page_source
    html = BeautifulSoup(raw, 'html.parser')
    time.sleep(15)
    container = html.select('div.review_card')
    checklist = list()
    total =  html.select_one('span.number').text.strip()
    for i in range(0,20):
         checklist.append(container[i])
    i = 0
    new = "Old"
    for con in checklist:
        
        t = base+"GB" #출처
        c = con.select_one("div.header>div>div>span.name").text.strip()
        pointlist = con.select('img')
        count = 0
        for point in pointlist:
            if(point.get('src') =="https://sf16-scmcdn-va.ibytedtos.com/obj/static-us/pico/developer_frontend/img/rating_star_yellow.0a718ebc.svg"):
                count += 1
        cou = "X"#글코멘트수
        h = con.select_one("div.header>div>div>span.time").text.strip() #글날짜
        l = con.select_one("div.content>div.review").text.strip() #글내용
        
        check = False
        if(i>=2):
            check = True
        if(new == "New"):
            new = SendAlarm(alarmList,i,l,check,t,h,str(count),1)
        else:
            new = SendAlarm(alarmList,i,l,check,t,h,str(count),1)
        if(new !="New"):
            i+=1
        
        worksheet.append_row([new,t,c,total,str(count), cou, h,l])# sheet 내 각 행에 데이터 추가
        time.sleep(1)
         
             
    # 중국 리뷰 
    time.sleep(10)
    FindButten(driver,By.XPATH,'//*[@id="tab-0"]')
    #but = driver.find_element(By.XPATH,value ='//*[@id="tab-0"]')
    #but.click()
    time.sleep(10)
    FindButten(driver,By.XPATH,'//*[@id="pane-0"]/div/div[2]/div[1]/div[1]/div[1]/div/span/span/i')
    #butRecent = driver.find_element(By.XPATH,value ='//*[@id="pane-0"]/div/div[2]/div[1]/div[1]/div[1]/div/span/span/i')
    #butRecent.click()
    time.sleep(10)
    FindButten2(driver,By.XPATH,'/html/body/div[3]/div[1]/div[1]/ul/li[2]')
    #butRecenttime = driver.find_element(By.XPATH,value ='/html/body/div[3]/div[1]/div[1]/ul/li[2]')
    #driver.execute_script("arguments[0].click()",butRecenttime)
    time.sleep(10)
    raw = driver.page_source
    html = BeautifulSoup(raw, 'html.parser')
    containerChina = html.select('div.review_card')
    
    FindButten(driver,By.XPATH,'//*[@id="pane-0"]/div/div[2]/div[1]/div[2]/div/div/span/span')
    time.sleep(10)
    #butShow = driver.find_element(By.XPATH,value ='//*[@id="pane-0"]/div/div[2]/div[1]/div[2]/div/div/span/span')
    #butShow.click()
    FindButten2(driver,By.XPATH,'/html/body/div[4]/div[1]/div[1]/ul/li[2]/span')
    #butEng = driver.find_element(By.XPATH,value ='/html/body/div[4]/div[1]/div[1]/ul/li[2]/span')
    #driver.execute_script("arguments[0].click()",butEng)
    time.sleep(10)
    raw = driver.page_source
    html = BeautifulSoup(raw, 'html.parser')
    containerEng = html.select('div.review_card')
    checklist.clear()
    for i in range(0,20):
         checklist.append(containerChina[i])
    index = 0
    i = 2
    total =  html.select_one('span.number').text.strip()
    for con in checklist:
        t = base+"China" #출처
        new ="Old"
          #print(con)
        c = con.select_one("div.header>div>div>span.name").text.strip() #글제목
        pointlist = con.select('img')
        count = 0
        for point in pointlist:
            if(point.get('src') =="https://sf16-scmcdn-va.ibytedtos.com/obj/static-us/pico/developer_frontend/img/rating_star_yellow.0a718ebc.svg"):
                count += 1
        cou = "X"#글코멘트수
        h = con.select_one("div.header>div>div>span.time").text.strip() #글날짜
        l = con.select_one("div.content>div.review").text.strip() #글내용
        eng = containerEng[index].select_one("div.content>div.review").text.strip() # 번역 내용
          
        check = False
        if(i>=4):
            check = True
        if(new == "New"):
            new = SendAlarm(alarmList,i-1,l,check, t,h,str(count),1,eng)
        else:
            new = SendAlarm(alarmList,i,l,check, t,h,str(count),1,eng)
        if(new !="New"):
            i+=1
          
        index += 1
        worksheet.append_row([new,t,c,total,str(count), cou, h,l,eng])# sheet 내 각 행에 데이터 추가
        time.sleep(1)
   
    #time.sleep(5)
    
    
    driver.close()


#CreateKakaoJson()

def WebCrawlingPicoGostop(urlLog,url1,url2,base):
    chrome_ver = chromedriver_autoinstaller.get_chrome_version()
    path=chromedriver_autoinstaller.install()
    driver = webdriver.Chrome()
    driver.get(urlLog)
    time.sleep(3)
    # 로그인하기
    #select=Select(driver.find_element(By.CSS_SELECTOR,value ="#review_context"))

    inputid = driver.find_element(By.XPATH,value ='//*[@id="root"]/main/div/div/div/article/article/article/form/div[1]/div[2]/input')
    inputpw = driver.find_element(By.XPATH,value ='//*[@id="root"]/main/div/div/div/article/article/article/form/div[2]/div/input')
    inputid.send_keys("howard@appnori.com")
    inputpw.send_keys("Appnori73")
    time.sleep(10)
    FindButten(driver,By.XPATH,'//*[@id="root"]/main/div/div/div/article/article/article/form/div[5]/button')
    #but = driver.find_element(By.XPATH,value ='//*[@id="root"]/main/div/div/div/article/article/article/form/div[5]/button')
    #but.click()
    time.sleep(10)
    FindButten(driver,By.XPATH,'//*[@id="dev-warp"]/div/div[3]/div[3]/div/div[1]/button')
    #butx = driver.find_element(By.XPATH,value ='//*[@id="dev-warp"]/div/div[3]/div[3]/div/div[1]/button')
    #butx.click()
    time.sleep(10)
    # pico 리뷰 페이지 
    #driver.get("https://developer-global.pico-interactive.com/console#/app/reviews/397/7098225807675359237")
    driver.get(url1)
    time.sleep(10)
    FindButten(driver,By.XPATH,'//*[@id="pane-1"]/div/div[2]/div[1]/div[1]/div[1]/div/span/span')
    #butRecent = driver.find_element(By.XPATH,value ='//*[@id="pane-1"]/div/div[2]/div[1]/div[1]/div[1]/div/span/span')
    #butRecent.click()
    FindButten2(driver,By.XPATH,'/html/body/div[3]/div[1]/div[1]/ul/li[3]/span')
    #butRecenttime = driver.find_element(By.XPATH,value ='/html/body/div[3]/div[1]/div[1]/ul/li[3]/span')
    #driver.execute_script("arguments[0].click()",butRecenttime)
    
    #butRecenttime = driver.find_element(By.XPATH,value ='/html/body/div[3]/div[1]/div[1]/ul')
    #driver.execute_script("arguments[0].click()",butRecenttime)
    #/html/body/div[5]/div[1]/div[1]/ul/li[2]
    time.sleep(10)
    raw = driver.page_source
    html = BeautifulSoup(raw, 'html.parser')
    time.sleep(10)
    container = html.select('div.review_card')
    checklist = list()
    total =  html.select_one('span.number').text.strip()
    for i in range(0,20):
         checklist.append(container[i])
    i = 0
    new = "Old"
    for con in checklist:
        
        t = base+"GB" #출처
        c = con.select_one("div.header>div>div>span.name").text.strip()
        pointlist = con.select('img')
        count = 0
        for point in pointlist:
            if(point.get('src') =="https://sf16-scmcdn-va.ibytedtos.com/obj/static-us/pico/developer_frontend/img/rating_star_yellow.0a718ebc.svg"):
                count += 1
        cou = "X"#글코멘트수
        h = con.select_one("div.header>div>div>span.time").text.strip() #글날짜
        l = con.select_one("div.content>div.review").text.strip() #글내용
        
        check = False
        if(i>=2):
            check = True
        if(new == "New"):
            new = SendAlarm(alarmList,i,l,check,t,h,str(count),2)
        else:
            new = SendAlarm(alarmList,i,l,check,t,h,str(count),2)
        if(new !="New"):
            i+=1
        
        worksheet.append_row([new,t,c,total,str(count), cou, h,l])# sheet 내 각 행에 데이터 추가
        time.sleep(1)

def checkNull(a):
    if(worksheet.cell(a+1, 7+1).value):
        alarmList.append(array[a].tolist()[7])
    else:
        alarmList.append("-1")

#SendTelegramBot(strnow)
#SendMsgForKakao(strnow)

#-*-coding:utf-8-*-
# 구글 스프레드 시트 연동 
scope = [
    'https://spreadsheets.google.com/feeds'
    ,'https://www.googleapis.com/auth/drive'
]

# 가상 스마트 메일 json 파일 
json_file_name = "C://Users//Administrator//Desktop//google_sheet//smart-amplifier-390002-d08fa835a9af.json"
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
gc = gspread.authorize(credentials)
print(gc)

# 연동 할려는 구글 스프레드 url
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1y0ipGFAf4j7ta-jHRzVMi5XYKQbaWjhrGZlhh9v894k/edit#gid=0'

# 스프레스시트 문서 가져오기
doc = gc.open_by_url(spreadsheet_url)
#print(doc)

index = 0
while True:
    #시트 선택하기
    worksheet = doc.worksheet('AIOReview')
    #worksheet = doc.worksheet('AIOReviewTest')
    worksheetBefore = doc.worksheet('AIOReviewBefore')


    array = np.array(worksheet.get_all_values())
    alarmList = list()
    checkNull(1)
    checkNull(2)
    checkNull(16)
    checkNull(17)
    checkNull(56)
    checkNull(57)
    checkNull(76)
    checkNull(77)
    checkNull(96)
    checkNull(101)

    #for row in alarmList:
    #    worksheetBefore.append_row([row]) 

    #worksheetBefore.clear()
    #for row in array:
    #    worksheetBefore.append_row(row.tolist())
    #    time.sleep(1)
    #print("이전 시트 저장 완료")

    #print(worksheet)
    #row_data = worksheet.row_values(1)
    #print(row_data)
    #range_list = worksheet.range('a1:m15')
    #print(range_list)


    worksheet.clear()
    #print("구글 스프레드 clear 되었습니다.")
    # 4. 데이터프레임 내 header(변수명)생성
    worksheet.append_row(["New","출처", "글제목","총평점","개인평점","글코멘트수", "글날짜","글내용","주석"])


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
        contenturl = requests.get(content)
        contenthtml = BeautifulSoup(contenturl.text, 'html.parser')
        contentcont = contenthtml.select_one("div.forum_op")
        contentcontcont = contentcont.select_one("div.content").text.strip()
        #print(contentcontcont)
        listcontent.append(contentcontcont)



    i=0
    new = "Old"
    for con in container:
        t = "SteamCommunity" #출처
        total = "x"
        point = "x"
        c = con.select_one("div.forum_topic_name").text.strip() #글제목
        cou = con.select_one("div.forum_topic_reply_count").text.strip() #글코멘트수
        h = con.select_one("div.forum_topic_lastpost").text.strip() #글날짜
        l = listcontent[i]
        check = False
        if(i>=2):
            check = True
        if(new == "New"):
            new = SendAlarm(alarmList,i,l,check,t,h,point)
        else:
            new = SendAlarm(alarmList,i,l,check,t,h,point)
        i+=1
        worksheet.append_row([new,t, c,total,point, cou, h,l])# sheet 내 각 행에 데이터 추가
    #print("steamcommunity가 구글 스프레드에 최신화 되었습니다.")




    WebCrawlingSteamReview("https://store.steampowered.com/app/1514840/allinone_sports_vr/","SteamReview",worksheet)
    #print("steamreview"+"가 구글 스프레드에 최신화 되었습니다.")

    WebCrawlingPico("https://sso-global.picoxr.com/passport?service=https%3a%2f%2fdeveloper-global.pico-interactive.com%2fconsole","https://developer-global.pico-interactive.com/console#/app/reviews/397/7098225807675359237","https://developer-global.pico-interactive.com/console#/app/reviews/397/2209","Pico",worksheet)
    #print("pico"+"가 구글 스프레드에 최신화 되었습니다.")



    WebCrawlingOculus2("https://www.oculus.com/experiences/quest/3840611616056575/?ranking_trace=0_3840611616056575_QUESTSEARCH_fcc9b3e7-dc82-4f7c-82f8-d90afedd0617","Oculus",worksheet)
    WebCrawlingOculus1("https://sidequestvr.com/app/4908/all-in-one-sports-vr","OculusSide",worksheet)


    #print("Oucule"+"가 구글 스프레드에 최신화 되었습니다.")
    worksheet.columns_auto_resize
    worksheet.rows_auto_resize
    #print("All-In_One구글 스프레드 사이즈 조절 완료")

    worksheet = doc.worksheet('SummerReview')
    #worksheet = doc.worksheet('AIOReviewTest')
    #worksheetBefore = doc.worksheet('SummerReviewBefore')
    cell_list = worksheet.findall("SteamCommunity")
    
    array = np.array(worksheet.get_all_values())
    alarmList = list()
    checkNull(1)
    checkNull(2)
    checkNull(21)
    checkNull(22)
    checkNull(41)
    checkNull(41+ len(cell_list))
    #worksheetBefore.clear()
    #for row in array:
    #    worksheetBefore.append_row(row.tolist())
    #    time.sleep(1)
    #print("이전 시트 저장 완료")
    worksheet.clear()
    #print("구글 스프레드 clear 되었습니다.")
    # 4. 데이터프레임 내 header(변수명)생성
    worksheet.append_row(["New","출처", "글제목","총평점","개인평점","글코멘트수", "글날짜","글내용","주석"])

    WebCrawlingPicoSummer("https://sso-global.picoxr.com/passport?service=https%3a%2f%2fdeveloper-global.pico-interactive.com%2fconsole","https://developer-global.pico-interactive.com/console#/app/reviews/397/7114478046865948678"," ","Pico")
    WebCrawlingSteamReviewCommu()
    WebCrawlingSteamReview("https://store.steampowered.com/app/2074750/AllInOne_Summer_Sports_VR/","SteamReview",worksheet)
    worksheet.columns_auto_resize
    worksheet.rows_auto_resize
    #print("Summer 구글 스프레드 사이즈 조절 완료")




    worksheet = doc.worksheet('RealGostop')
    #worksheetBefore = doc.worksheet('RealGostopBefore')


    array = np.array(worksheet.get_all_values())
    alarmList = list()
    checkNull(1)
    checkNull(2)
    

    #worksheetBefore.clear()
    #for row in array:
    #    worksheetBefore.append_row(row.tolist())
    #    time.sleep(1)
    #print("이전 시트 저장 완료")

    worksheet.clear()
    #print("구글 스프레드 clear 되었습니다.")
    # 4. 데이터프레임 내 header(변수명)생성
    worksheet.append_row(["New","출처", "글제목","총평점","개인평점","글코멘트수", "글날짜","글내용","주석"])
    WebCrawlingPicoGostop("https://sso-global.picoxr.com/passport?service=https%3a%2f%2fdeveloper-global.pico-interactive.com%2fconsole","https://developer-global.pico-interactive.com/console#/app/reviews/397/7205734342381961221"," ","Pico")
    worksheet.columns_auto_resize
    worksheet.rows_auto_resize
    index += 1
    print(str(index)+"회차 구글 스프레드 수정 완료")
    waitingTime = 3600
    for i in range(0,waitingTime,1):
        time.sleep(1)
    
"""
from datetime import datetime

now = datetime.now()
strnow = now.strftime("%Y %m %d %H %M %S.%f")
print(strnow)
with open(r"time.txt", "a") as fp:L
      fp.write(str(strnow)+'\n')
"""