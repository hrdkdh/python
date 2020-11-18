import re
import sys
import urllib
import requests
import pandas as pd
from time import sleep
from datetime import datetime
from bs4 import BeautifulSoup as bs

#로그인 아이디/비번을 매번 입력하지 않으려면 setData() 함수로 이동하여 값을 미리 입력해 놓으세요

login_id = None
login_pw = None
from_date = None
to_date = None
search_title = None
login_data = None
survey_data_excel = None
survey_data_voc = None
login_url = "https://e-campus.posco.co.kr/UserMain/portal_loginTop.jsp"
now_datetime = str(int(datetime.now().timestamp()))

def getExcelData():
    print("이캠퍼스 접속중...")
    with requests.Session() as s:
        login_req = s.post(login_url, data=login_data)
        if login_req.status_code != 200:
            print(login_req.status_code)
            print("관리자 화면 로그인에 실패하였습니다.")
        else:
            survey_page = s.post("https://e-campus.posco.co.kr/AttendingMgr/S200406401.jsp", data=survey_data_excel)
            download_file_name = "C:\\Users\\POSCOUSER\\Desktop\\이캠퍼스_설문점수 다운로드_" + now_datetime + ".xls"
            try:
                open(download_file_name, "wb").write(survey_page.content)
                print("파일을 다운로드하였습니다. 저장경로 : " + download_file_name)
            except:
                print("파일 다운로드에 실패하였습니다.")

def getVocData():
    print("이캠퍼스 접속중...")
    with requests.Session() as s:
        login_req = s.post(login_url, data=login_data)
        if login_req.status_code != 200:
            print(login_req.status_code)
            print("관리자 화면 로그인에 실패하였습니다.")
        else:
            print("VOC 데이터를 정리하는 중...")
            results_data = []
            survey_page = s.post("https://e-campus.posco.co.kr/AttendingMgr/S200406400.jsp", data=survey_data_voc)
            download_file_name = "C:\\Users\\POSCOUSER\\Desktop\\이캠퍼스_VOC 다운로드_" + now_datetime + ".xlsx"
            list_soup = bs(survey_page.content, "html.parser")
            trs = list_soup.findAll("table")[0].findAll("tbody")[0].findAll("tr")
            for tr in trs:
                this_cha_name = tr.findAll("td")[0].text
                this_survey_name = tr.findAll("td")[1].text
                this_start_date = tr.findAll("td")[2].text
                this_end_date = tr.findAll("td")[3].text
                this_results_link = tr.findAll("td")[0].findAll("a")[0].attrs["href"]
                this_results_real_link = "https://e-campus.posco.co.kr/AttendingMgr/S200406410.jsp?CLC_E_PROJECT_ID=" + this_results_link.split("(")[1].split(",")[0] + "&CLC_E_PAPER_ID=" + this_results_link.split("(")[1].split(",")[1]
                result_page = s.get(this_results_real_link)
                result_soup = bs(result_page.content, "html.parser")
                result_trs_temp = result_soup.findAll("table")[0].findAll("tbody")[0]
                result_trs = str(result_trs_temp).replace("</tr>", "").replace("<tbody>", "").replace("</tbody>", "").replace("<tr>", "</tr><tr>")[6:].strip().split("</tr>")
                for result_tr in result_trs:
                    if len(result_tr) > 10 and len(result_tr.split("</td>")) > 0:
                        this_tr_content = result_tr.split("</td>")
                        for i, result_td in enumerate(this_tr_content):
                            cleanr =re.compile('<.*?>')
                            result_td_text = re.sub(cleanr, "", result_td).strip()
                            if i > 1 and result_td_text.isdigit() is False and len(result_td_text) > 1:
                                results_data.append({
                                    "차수명" : this_cha_name,
                                    "설문명" : this_survey_name,
                                    "시작일" : this_start_date,
                                    "종료일" : this_end_date,
                                    "응답자" : re.sub(cleanr, "", this_tr_content[0]),
                                    "직책" : re.sub(cleanr, "", this_tr_content[1]),
                                    "문항번호" : str(i+1),
                                    "응답내용" : result_td_text
                                })
            try:
                df = pd.DataFrame(results_data)
                df.to_excel(download_file_name,
                    sheet_name = "VOC결과",
                    header = True,
                    index = True,
                    index_label = "id", 
                    startrow = 0, 
                    startcol = 0, 
                    #engine = 'xlsxwriter'
                )
                print("VOC를 저장하였습니다. 저장경로 : " + download_file_name)
            except:
                print("VOC 저장에 실패하였습니다.")

def setData():
    global login_id, login_pw, from_date, to_date, search_title, login_data, survey_data_excel, survey_data_voc

    #로그인 아이디/비번을 매번 입력하지 않으려면 아래 두줄 주석처리할 것
    login_id = input("이캠퍼스 아이디를 입력해 주세요 : ")
    login_pw = input("이캠퍼스 패스워드를 입력해 주세요 : ")

    #로그인 아이디/비번을 매번 입력하지 않으려면 아래 두줄 주석해제하고 아이디/비번을 미리 입력해 놓을 것
    # login_id = ""
    # login_pw = ""

    from_date = input("통계 검색 시작일을 YYYYMMDD 형태로 입력해 주세요(오늘 날짜로 하려면 그냥 엔터) : ")
    to_date = input("통계 검색 종료일을 YYYYMMDD 형태로 입력해 주세요(오늘 날짜로 하려면 그냥 엔터) : ")
    search_title = input("통계 검색어를 입력해 주세요(전체 검색하려면 그냥 엔터) : ")

    if login_id == "" or login_pw == "":
        print("                                                              ")
        print("아이디/패스워드를 입력하지 않았습니다!")
        setData()
    else:
        login_data = {
            "userid": login_id,
            "password": login_pw,
            "portal" : "portal"
        }

        from_date = checkYmdAndChangeFormat(from_date)
        to_date = checkYmdAndChangeFormat(to_date)

        if from_date == None or to_date == None:
            setData()
        else:
            survey_data_excel = {
                "CLC_EDU_CYBER_YN": "N",
                "FROM_DATE": from_date,
                "TO_DATE": to_date,
                "TITLE": search_title,
                "PAPER_KEYWORD":
                ""
            }
            survey_data_voc = {
                "PAGE_S200406400": "1",
                "CLC_EDU_CYBER_YN": "N",
                "FROM_DATE": from_date,
                "TO_DATE": to_date,
                "TITLE": search_title,
                "PAPER_KEYWORD":
                ""
            }

def checkYmdAndChangeFormat(ymd):
    if len(ymd) == 8: #검색일을 8자리 숫자로 정확하게 입력하였다면 중간에 bar(-)를 넣어줌
        ymd = ymd[0:4] + "-" + ymd[4:6] + "-" + ymd[6:8]
    elif len(ymd) == 0: #검색일을 넣지 않았다면 현재 날짜로 넣어줌
        ymd = datetime.today().strftime("%Y-%m-%d")
    else: #검색일을 넣었지만 8자리로 넣지 않았다면 None으로 출력
        print("                                                              ")
        print("통계 검색일을 정확히 입력하지 않았습니다.")
        ymd = None
    return ymd

def selectFunc():
    print("===================================================================================================") 
    print("1 : 설문점수 엑셀로 다운로드")
    print("2 : VOC 다운로드")
    print("3 : 프로그램 종료")    
    print("===================================================================================================") 
    func = input("사용할 기능의 번호를 입력한 후 엔터키를 눌러주세요.")
    if func not in ["1", "2", "3"]:
        print("                                                              ")
        print("정확한 번호를 입력해 주세요.")
        print("                                                              ")
        sleep(2)
        selectFunc()
    else:
        if func == "1":
            setData()
            print("설문점수를 엑셀로 다운로드 받습니다.")
            getExcelData()
            sleep(2)
            selectFunc()
        elif func == "2":
            setData()
            print("VOC를 엑셀로 다운로드 받습니다.")
            getVocData()
            sleep(2)
            selectFunc()
        elif func == "3":
            print("프로그램을 종료합니다.")
            sys.exit()

if __name__ == "__main__":
    selectFunc()