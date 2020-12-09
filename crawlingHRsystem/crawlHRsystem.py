import sys
import json
import urllib
import requests
import warnings
import pandas as pd
from bs4 import BeautifulSoup as bs
from datetime import datetime

warnings.filterwarnings('ignore', message='Unverified HTTPS request')

login_id = ""
login_pw = ""
emp_no = ""
emp_nm = ""

def getYmPeriod(from_ym, to_ym):
    result_arr = []
    from_year = int(from_ym[0:4])
    to_year = int(to_ym[0:4])
    from_month = int(from_ym[4:6])
    to_month = int(to_ym[4:6])

    for year in range(from_year, to_year+1):
        for month in range(1, 13):
            this_month = str(month)
            if len(this_month) < 2:
                this_month = "0" + this_month
            this_ym = str(year) + this_month
            
            if year == from_year and year != to_year:
                if month >= from_month:
                    result_arr.append(this_ym)
            elif year == from_year and year == to_year:
                if month >= from_month and month <= to_month:
                    result_arr.append(this_ym)
            elif year > from_year and year < to_year: 
                result_arr.append(this_ym)
            elif year != from_year and year == to_year:
                if month <= to_month:
                    result_arr.append(this_ym)
    return result_arr

def crawlData(from_ym, to_ym):
    print("크롤링중...")
    ym_list = getYmPeriod(from_ym, to_ym)
    first_url = "http://swpsso.posco.net/idms/U61/jsp/login/login.jsp"
    login_url = "http://swpsso.posco.net/idms/U61/jsp/login/loginProc.jsp"
    session_url1 = "http://swpsso.posco.net/idms/U61/jsp/manysession.jsp"
    session_url2 = "http://swpsso.posco.net/pkmsdisplace"
    main_page_url = "http://swp.posco.net/wps/index.jsp"
    sso_page_url = "https://hr.poscohrd.com/sso_login.jsp"
    hr_menu_url = "https://hr.poscohrd.com/menu.jsp?menuId=PVT_M_HRM&isPrivateMenu=null"
    hr_h5_url = "https://hr.poscohrd.com/serviceBroker.h5"
    login_data = {
        "username": login_id,
        "password": login_pw,
        "login-form-type": "pwd"
    }
    headers = {
        "Accept" : "text/html, application/xhtml+xml, image/jxr, */*",
        "Accept-Encoding" : "gzip, deflate",
        "Accept-Language" : "ko",
        "Cache-Control" : "no-cache",
        "Connection" : "Keep-Alive",
        "Content-Type" : "application/x-www-form-urlencoded",
        "Host" : "swpsso.posco.net",
        "Referer" : "http://swpsso.posco.net/idms/U61/jsp/login/login.jsp",
        "User-Agent" : "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; TCO_20201207074412; rv:11.0) like Gecko"
    }
    with requests.Session() as s:    
        s.get(first_url, headers = headers)
        s.post(login_url, headers = headers, data = login_data)
        s.get(session_url1, headers = headers)
        s.get(session_url2)
        s.get(main_page_url)

        requests.packages.urllib3.util.ssl_.DEFAULT_CIPHERS += 'HIGH:!DH:!aNULL'
        try:
            requests.packages.urllib3.contrib.pyopenssl.DEFAULT_SSL_CIPHER_LIST += 'HIGH:!DH:!aNULL'
        except AttributeError:
            pass
        
        hr_connect_data = {
            "ssoToken" : s.cookies["SWP-H-SESSION-ID"],
            "websealType" : "I",
            "domainName" : "swpsso.posco.net",
            "serverName" : "default2-webseald-PLCSSOA6",
            "LANG" : "ko",
            "TIMEZONE" : "+540"
        }
        s.post(sso_page_url, data = hr_connect_data, verify=False)
        hr_menu_req = s.get(hr_menu_url, verify=False)
        sessionId = s.cookies.get(name="JSESSIONID", domain="hr.poscohrd.com", path="/")
        sessionId_for_h5 = hr_menu_req.text.split("_sessionId = \"")[1].split("\";")[0]
        header_for_payment = {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
            "Accept-Encoding": "gzip, deflate, br",
            "Accept-Language": "ko,ko-KR;q=0.9,en;q=0.8",
            "Cache-Control": "max-age=0",
            "Connection": "keep-alive",
            "Content-Length": "429",
            "Content-Type": "application/x-www-form-urlencoded",
            "Cookie": "JSESSIONID="+sessionId,
            "Host": "hr.poscohrd.com",
            "Origin": "null",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "none",
            "Sec-Fetch-User": "?1",
            "Upgrade-Insecure-Requests": "1",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36"
        }
        data_for_payment = {
            "request_message":"{\"HEADER\":{\"companyCd\":\"01\",\"serviceId\":\"FRM_OPEN_OBJECT_THIN\",\"sessionId\":\""+ sessionId_for_h5 +"\",\"localeCd\":\"KO\",\"langCd\":\"KO\"},\"BODY\":{\"ME_OBJECT_REQUEST\":[{\"objectId\":\"PAY0025_98\",\"param\":\"\",\"date\":\"\"}]}}"
        }
        hr_payment_req = s.post(hr_h5_url, data = data_for_payment, verify=False)
        emp_id = hr_payment_req.text.split(", 'emp_id': '")[1].split("' }; };")[0]
        df = None
        count = 0
        for ym in ym_list:
            year = ym[0:4]
            month = ym[4:6]
            for pay_ymd_id in range(1,11):
                payload_for_payment = {
                    "HEADER":
                    {
                        "companyCd":"01",
                        "sessionId":sessionId_for_h5,
                        "serviceId":"PAY0025_00_R01_99",
                        "objectId":"PAY0025_98",
                        "actionType":"retrieve",
                        "localeCd":"KO",
                        "langCd":"KO"
                    },
                    "BODY":
                    {
                        "ME_PAY0025_01":
                        [{
                            "_seq":"",
                            "sStatus":"U",
                            "sDelete":"",
                            "company_cd":"01",
                            "locale_cd":"KO",
                            "pay_ymd_id":pay_ymd_id,
                            "emp_id":emp_id,
                            "auth_str":"admin",
                            "emp_no":emp_no,
                            "emp_nm":emp_nm,
                            "retireyn":"Y",
                            "pcheck":"Y",
                            "pay_year":year,
                            "pay_month":month,
                            "pay_ym":ym
                        }]
                    }
                }
                header_for_hr = {
                    "Accept": "application/json, text/javascript, */*; q=0.01",
                    "Accept-Encoding": "gzip, deflate, br",
                    "Accept-Language": "ko,ko-KR;q=0.9,en;q=0.8",
                    "Connection": "keep-alive",
                    "Content-Length": "328",
                    "Content-Type": "application/json; charset=UTF-8",
                    "Cookie": "JSESSIONID="+sessionId,
                    "Host": "hr.poscohrd.com",
                    "Origin": "https://hr.poscohrd.com",
                    "Referer": "https://hr.poscohrd.com/menu.jsp?menuId=PVT_M_HRM&isPrivateMenu=null",
                    "Sec-Fetch-Dest": "empty",
                    "Sec-Fetch-Mode": "cors",
                    "Sec-Fetch-Site": "same-origin",
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36",
                    "X-Requested-With": "XMLHttpRequest"
                }
                results_src = s.post(hr_h5_url, headers=header_for_hr, data=json.dumps(payload_for_payment), verify=False)
                results = json.loads(results_src.text)
                val_check = results.get("BODY").get("ME_PAY0025_04")[0].get("d_amt01")
                if len(val_check) > 0:
                    count += 1
                    results_body = results.get("BODY")
                    sum_data = results_body.get("ME_PAY0025_06")[0]
                    result_arr = [
                        {"get_name" : "ME_PAY0025_03"},
                        {"get_name" : "ME_PAY0025_04"}
                    ]
                    columns = ["년월", "급여총액", "공제총액", "실지급액"]
                    rows = [ym, sum_data["totpayamt"], sum_data["totdeducamt"], sum_data["actpayamt"]]
                    outcome_dict = {}
                    for idx, result_item in enumerate(result_arr):
                        if idx == 0:
                            pre = "a"
                        elif idx == 1:
                            pre = "d"
                        for key, val in results_body.get(result_item["get_name"])[0].items():
                            if key[0:4] == pre+"_nm":
                                if val == "기준연봉월할액" or val == "경영성과금" or val =="업적연봉" or val == "격려금":
                                    val = "기준연봉월할액/경영성과금/업적연봉/격려금"
                                columns.append(val)
                                outcome_dict[val] = key[4:6]
                            if key[0:5] == pre+"_amt":
                                for outcome_key, outcome_val in outcome_dict.items():
                                    if outcome_val == key[5:7]:
                                        rows.append(val)
                                        outcome_dict[outcome_key] = val
                                        break
                    if df is None:
                        df = pd.DataFrame([rows], columns=columns)
                    else:
                        # this_df = pd.DataFrame([rows], columns=columns)
                        df = pd.concat([df, pd.DataFrame([rows], columns=columns)], ignore_index = True)
        print("총 카운트 : " + str(count) + "건")
        # print(df)
        file_name = "output_" + from_ym + "-" + to_ym + ".xlsx"
        df.to_excel(
            file_name,
            # sheet_name = result_item["sheet_name"],
            header = True,
            index = True,
            index_label = "id",
            startrow = 0, 
            startcol = 0
        )
    print("크롤링을 완료하였습니다. " + file_name + "로 저장하였습니다.")

if __name__ == "__main__":
    login_id = input("아이디를 입력해 주십시오 : ")
    login_pw = input("비밀번호를 입력해 주십시오 : ")
    emp_no = input("직번을 입력해 주십시오 : ")
    emp_nm = input("이름를 입력해 주십시오 : ")
    from_ym = input("시작년월을 입력해 주십시오(YYYYMM) : ")
    to_ym = input("종료년월을 입력해 주십시오(YYYYMM) : ")
    crawlData(from_ym, to_ym)