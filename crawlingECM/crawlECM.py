import sys
import json
import urllib
import requests
import pandas as pd
from bs4 import BeautifulSoup as bs

login_id = ""
login_pw = ""

def makeUrlParam(formId="F_LIST_MYDEPT", listKey="cab0000bf4b95ae60c4_cab0000bf4b95ae60c4", docCntPerPage="500", fileType="ALL_FORMAT"):
    query = [
        ("ServiceName", "DocList-service"),
        ("getDocList", "true"),
        ("isHTML", "T"),
        ("formId", formId),
        ("listKey", listKey),
        ("searchValue", ""),
        ("pageNumber", "1"),
        ("docCntPerPage", docCntPerPage),
        ("searchPeriod", "ALL"),
        ("wfStatusType", ""),
        ("fileType", "ALL_FORMAT"),
        ("orderBy", ""),
        ("sortDirection", "desc"),
        ("selCabinetType", ""),
        ("logMenuId", "")
    ]
    result = urllib.parse.urlencode(query, doseq=True)
    return result    

def getFormDataFromHtml(html):
    soup = bs(html, "html.parser")
    arr = soup.select("input")
    results = {}
    for item in arr:
        if "name" in item.attrs and len(item.attrs["name"]) > 0:
            results[item.attrs["name"]] = item.attrs["value"]
    return results

def crawlData(ecm_ajax_url):
    first_url = "http://swpsso.posco.net/idms/U61/jsp/login/login.jsp"
    login_url = "http://swpsso.posco.net/idms/U61/jsp/login/loginProc.jsp"
    session_url1 = "http://swpsso.posco.net/idms/U61/jsp/manysession.jsp"
    session_url2 = "http://swpsso.posco.net/pkmsdisplace"
    main_page_url = "http://swp.posco.net/wps/index.jsp"
    sso_page_url = "http://swpsso.posco.net/idms/U61/jsp/redirectSMSP.jsp?redir_url=http%3A%2F%2Fswpecm.posco.net%3A7091%2FECM%2Findex.jsp"
    ecm_page_url = "http://swpecm.posco.net:7091/ECM/index.jsp"
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
    headers_for_ecm_ajax = {
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "ko,ko-KR;q=0.9,en;q=0.8",
        "Connection": "keep-alive",
        "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
        "Host": "swpecm.posco.net:7091",
        "Referer": "http://swpecm.posco.net:7091/ECM/jsp/main/ecmMain.jsp",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36",
        "X-Requested-With": "XMLHttpRequest"
    }
    with requests.Session() as s:    
        s.get(first_url, headers = headers)
        s.post(login_url, headers = headers, data = login_data)
        s.get(session_url1, headers = headers)
        s.get(session_url2)
        s.get(main_page_url)
        sso_data = s.get(sso_page_url)
        form_data_for_ecm_login = getFormDataFromHtml(sso_data.text)
        redrt_src = s.post(ecm_page_url, data = form_data_for_ecm_login)
        redrt = "http://swpecm.posco.net:7091/ECM" + redrt_src.text.split(" actionUrl = '")[1].replace("'</script>", "")
        s.get(redrt)
        results = json.loads(s.get(ecm_ajax_url, headers = headers_for_ecm_ajax).text)
        columns = ["폴더명", "파일명", "소유자", "등급", "공개여부", "상태", "등록일시", "수정일시"]
        rows = []
        for doc in results.get("ML_DOC_LIST"):
            rows.append([doc.get("MS_CABINET_NAME"), doc.get("MS_OBJECT_NAME"), doc.get("MS_OWNER_NAME"), doc.get("MS_SECURITY_LEVEL_TEXT"), doc.get("MS_OPEN_FLAG_TEXT"), doc.get("MS_STATUS_TEXT"), doc.get("MS_REG_DATE"), doc.get("MS_FILE_MODIFY_DATE")]) 
        df = pd.DataFrame(rows, columns=columns)
        df.to_excel(
            "output.xlsx",
            header = True,
            index = True,
            startrow = 0, 
            startcol = 0
        )
    print("크롤링을 완료하였습니다. output.xlsx 파일로 저장하였습니다.")

if __name__ == "__main__":
    login_id = input("아이디를 입력해 주십시오 : ")
    login_pw = input("비밀번호를 입력해 주십시오 : ")
    url_param = makeUrlParam("F_LIST_MYDEPT", "cab0000bf4b95ae60c4_cab0000bf4b95ae60c4", "2000", "ALL_FORMAT")
    ecm_ajax_url = "http://swpecm.posco.net:7091/ECM/ajaxAction.do?"+url_param
    crawlData(ecm_ajax_url)