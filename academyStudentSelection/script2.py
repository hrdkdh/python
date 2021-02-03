"""
----- 2nd script ----

"""

import requests
import pandas as pd
from time import sleep
from urllib.parse import urlencode
from bs4 import BeautifulSoup as bs

def set_app_data(cha_name, result_df):
    login_id = input("취창업캠프사이트 관리자 아이디를 입력해 주세요 : ")
    login_pw = input("취창업캠프사이트 관리자 패스워드를 입력해 주세요 : ")

    print("youth.posco.com 접속...")
    base_url = "http://youth.posco.com/posco/_owner/"
    login_url = base_url+"index.php?act=login"
    login_data = {
        "wd_id": login_id,
        "wd_pw": login_pw
    }
    with requests.Session() as s:
        print("youth.posco.com 로그인...")
        login_req = s.post(login_url, data=login_data)
        if login_req.status_code != 200:
            print(login_req.status_code)
            print("관리자 화면 로그인에 실패하였습니다.")
            sleep(2)
            set_app_data(cha_name, result_df)
        
        #과정명 검색결과 출력 + 제대로 로그인되었는지 체크
        cha_list_data = s.get(base_url+"index.php?mod=lecture&act=main&cate=&sField=&sValue="+cha_name)
        soup = bs(cha_list_data.text, "html.parser")
        table = soup.select("table")
        try:
            strongs = table[1].select("strong")
            print("youth.posco.com 로그인 성공")
        except Exception as e:
            print("youth.posco.com 아이디/비번을 잘못 입력하였습니다.", e)
            sleep(2)
            set_app_data(cha_name, result_df)

        student_page_url = ""
        for strong in strongs:
            if "수 : " in strong.get_text() and "접" in strong.get_text()[:1]:
                href = strong.parent.attrs["href"]
                student_page_url=href[2:len(href)]
                break

        if student_page_url == "":
            print("차수명이 잘못되어 교육생 정보 업데이트에 실패하였습니다.")
            print("차수명을 정확히 입력한 후 다시 시도해 주세요.")
            exit()

        student_data = s.get(base_url+student_page_url)
        student_data_soup = bs(student_data.text, "html.parser")
        student_data_table = student_data_soup.select(".ay_table")[0].select("tr")
        count = 0
        update_list_total = []
        update_list_passed = []
        update_list_unpassed = []
        update_list_waited = []
        for tr in student_data_table:
            if len(tr.select("td")) > 10 and len(tr.select("td")[0].select("input")) > 0 and tr.select("td")[0].select("input")[0].attrs["name"] == "luCode[]":
                this_update_dict = {}
                this_update_dict["std_name"] = tr.select("td")[3].select("strong")[0].text
                this_update_dict["std_phone"] = tr.select("td")[4].select("strong")[0].text
                this_update_dict["id_no"] = tr.select("td")[0].select("input")[0].attrs["value"]
                list_append_check = False
                for i in range(len(result_df["휴대폰"])):
                    if result_df["합격"].loc[i] == "합격" or result_df["합격"].loc[i] == "불합격" or result_df["합격"].loc[i] == "대기":
                        if result_df["휴대폰"].loc[i] == this_update_dict["std_phone"]:
                            this_update_dict["result"] = result_df["합격"].loc[i]
                            list_append_check = True
                            break
                if list_append_check:
                    update_list_total.append(this_update_dict)
                    if this_update_dict["result"] == "합격":
                        update_list_passed.append(this_update_dict)
                    if this_update_dict["result"] == "불합격":
                        update_list_unpassed.append(this_update_dict)
                    if this_update_dict["result"] == "대기":
                        update_list_waited.append(this_update_dict)
        
        update_form_info = {
            "actType" : "",
            "grant" : "",
            "part" : "",
            "lgCode" : "",
            "leCode" : "",
        }
        for inp in student_data_soup.select("input"):
            for input_name, _ in update_form_info.items():
                if inp.attrs["name"] == input_name:
                    update_form_info[input_name] = inp.attrs["value"]
                    break
        
        #합격자 업데이트
        update_std_status(s, base_url, update_list_passed, update_form_info, "합격")
        #불합격자 업데이트
        update_std_status(s, base_url, update_list_unpassed, update_form_info, "불합격")
        #대기자 업데이트
        update_std_status(s, base_url, update_list_waited, update_form_info, "대기자")

def update_std_status(s, base_url, update_list, update_form_info, ststusSel):
    if len(update_list) > 0:
        this_update_form_info = update_form_info
        this_update_form_info["ststusSel"] = ststusSel
        id_no_list = []
        for std in update_list:
            id_no_list.append(std["id_no"])
        this_update_form_info["luCode[]"] = id_no_list
        result = s.post(base_url+"index.php?mod=lecture&act=dataRegOk", data = this_update_form_info)
        if "opener.location.reload();" in result.text:
            print("----------------------------")
            print("업데이트 완료({}) : 총 {}명".format(ststusSel, len(update_list)))
            print("업데이트 명단({})".format(ststusSel), end=" : ")
            for std in update_list:
                print(std["std_name"], end=" ")
            print("\n")

if __name__ == "__main__":
    try:
        result_df = pd.read_clipboard()
    except:
        print("엑셀 파일에서 전체셀을 선택, 복사한 후 다시 실행해 주세요.")
    finally:
        if len(result_df)>0:
            print("복사된 엑셀 데이터를 로드 완료하였습니다.")
            cha_name = input("과정명을 입력하세요: ")
            set_app_data(cha_name, result_df)
            print("합격/불합격/대기자 업데이트가 완료되었습니다.")
        else:
            print("엑셀 파일에서 전체셀을 선택, 복사한 후 다시 실행해 주세요.")