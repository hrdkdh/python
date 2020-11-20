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
headers = {
    "Accept" : "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
    "Accept-Encoding" : "gzip, deflate, br",
    "Accept-Language" : "ko,ko-KR;q=0.9,en;q=0.8",
    "Cache-Control" : "max-age=0",
    "Connection" : "keep-alive",
    # "Content-Length" : "46",
    "Content-Type" : "application/x-www-form-urlencoded",
    # "Cookie" : "_ga=GA1.3.1808224504.1589257481; gitple-m-1merE0mUs6OXn3H7MMdpUUvIm5SqWF80={"state":"close"}; JSESSIONID=s0XjN0rf2Ctd__GW2k2Lkkw7bJlZqNNiKaSqh0StfM3rJ1WuFUn4!-1721597082",
    "Host" : "e-campus.posco.co.kr",
    "Origin" : "http://e-campus.posco.co.kr",
    "Referer" : "http://e-campus.posco.co.kr/",
    "Sec-Fetch-Dest" : "document",
    "Sec-Fetch-Mode" : "navigate",
    "Sec-Fetch-Site" : "cross-site",
    "Sec-Fetch-User" : "?1",
    "Upgrade-Insecure-Requests" : "1",
    "User-Agent" : "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.193 Safari/537.36"
}

def getExcelData():
    print("이캠퍼스 접속중...")
    now_datetime = str(int(datetime.now().timestamp()))
    with requests.Session() as s:
        login_req = s.post(login_url, headers = headers, data=login_data)
        if login_req.status_code != 200:
            print(login_req.status_code)
            print("관리자 화면 로그인에 실패하였습니다.")
        else:
            survey_page = s.post("https://e-campus.posco.co.kr/AttendingMgr/S200406401.jsp", headers = headers, data=survey_data_excel)
            download_file_name = "이캠퍼스_설문점수 다운로드 결과_" + now_datetime + ".xls"
            try:
                open(download_file_name, "wb").write(survey_page.content)
                print("파일을 다운로드하였습니다. 저장경로 : " + download_file_name)
            except:
                print("파일 다운로드에 실패하였습니다.")

def getVocData():
    print("이캠퍼스 접속중...")
    now_datetime = str(int(datetime.now().timestamp()))
    with requests.Session() as s:
        login_req = s.post(login_url, headers = headers, data=login_data)
        if login_req.status_code != 200:
            print(login_req.status_code)
            print(login_req.text)
            print("관리자 화면 로그인에 실패하였습니다.")
        else:
            print("VOC 데이터를 정리하는 중...")
            results_data = []
            survey_page = s.post("https://e-campus.posco.co.kr/AttendingMgr/S200406400.jsp", headers = headers, data=survey_data_voc)
            download_file_name = "이캠퍼스_VOC 다운로드 결과_" + now_datetime + ".xlsx"
            page_soup = bs(survey_page.content, "html.parser")
            pages = page_soup.select(".paginate")[0].findAll("a")

            for page in range(1, len(pages)+2): #페이지별 loop
                print("  ")
                print("총 " + str(len(pages)+1) + "페이지 중 " + str(page) + "페이지 크롤링중...")
                survey_data_voc["PAGE_S200406400"] = page
                this_survey_page = s.post("https://e-campus.posco.co.kr/AttendingMgr/S200406400.jsp", headers = headers, data=survey_data_voc)
                
                list_soup = bs(this_survey_page.content, "html.parser")
                trs = list_soup.findAll("table")[0].findAll("tbody")[0].findAll("tr")
                now_results_data_count = len(results_data)
                for tr in trs: #설문별 loop
                    this_cha_name = tr.findAll("td")[0].text
                    this_survey_name = tr.findAll("td")[1].text
                    this_start_date = tr.findAll("td")[2].text
                    this_end_date = tr.findAll("td")[3].text
                    this_results_link = tr.findAll("td")[0].findAll("a")[0].attrs["href"]
                    this_results_real_link = "https://e-campus.posco.co.kr/AttendingMgr/S200406410.jsp?CLC_E_PROJECT_ID=" + this_results_link.split("(")[1].split(",")[0] + "&CLC_E_PAPER_ID=" + this_results_link.split("(")[1].split(",")[1]
                    result_page = s.get(this_results_real_link, headers = headers)
                    result_soup = bs(result_page.content, "html.parser")
                    result_trs_temp = result_soup.findAll("table")[0].findAll("tbody")[0]
                    result_trs = str(result_trs_temp).replace("</tr>", "").replace("<tbody>", "").replace("</tbody>", "").replace("<tr>", "</tr><tr>")[6:].strip().split("</tr>")
                    for result_tr in result_trs: #설문응답자별 loop
                        if len(result_tr) > 10 and len(result_tr.split("</td>")) > 0:
                            this_tr_content = result_tr.split("</td>")
                            for i, result_td in enumerate(this_tr_content): #설문응답 내용별 loop
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
                print(" ┗ VOC " + str(len(results_data)-now_results_data_count) + "건 입력완료(누적 " + str(len(results_data)) + "건)")
            try:
                df = pd.DataFrame(results_data)
                df.to_excel(download_file_name,
                    sheet_name = "VOC결과",
                    header = True,
                    index = True,
                    index_label = "id", 
                    startrow = 0, 
                    startcol = 0
                )
                print("  ")
                print("VOC를 저장하였습니다. 저장경로 : " + download_file_name)
            except:
                print("  ")
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
    print("이캠퍼스 집합과정 설문 결과를 손쉽게 다운받을 수 있는 도구입니다.")
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