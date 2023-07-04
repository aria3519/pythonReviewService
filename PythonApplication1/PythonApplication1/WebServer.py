
import requests
import json
from flask import Flask,request,jsonify

application = Flask(__name__)
@application.route('/')


def KakaoMessage():
    req = request.get_json()
    text ="111111"

    res = {"version:":"2.0",
           "template":{
               "outputs:":[
                   {"simpleText":{"text":request.remote_addr}}]}}

    return jsonify(res)

"""
def Message():
    content = request.get_json()
    content = content['userRequest']['utterance']
    content=content.replace("\n","")
    print(content)
    if content == u"오늘의 메뉴":
        dataSend = {
            "version" : "2.0",
            "template" : {
                "outputs" : [
                    {
                        "simpleText" : {
                            "text" : "테스트입니다."
                        }
                    }
                ]
            }
        }
    else:
        dataSend = {
            "version" : "2.0",
            "template" : {
                "outputs" : [
                    {
                        "simpleText" : {
                            "text" : "error입니다."
                        }
                    }
                ]
            }
        }
    return jsonify(dataSend)
"""
application.run()

if __name__ == "__main__" :
        application.run(host='0.0.0.0',port = 5000,threaded =Ture)

