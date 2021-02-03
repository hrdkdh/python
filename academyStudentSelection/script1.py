"""
----- 1st script ----
1) 접수자 엑셀 다운로드
2) 계산식에 의해 (가)합격자/(가)대기자/(가)불합격자 판별하여 가합격-가대기-가불합 순으로 정렬, 엑셀로 저장
3) 엑셀 가합격 셀에 합격여부 표기하여 저장
3) 지원동기 엑셀 다운로드
4) 가합격자 엑셀에 개인별로 지원동기 셀을 추가하여 종합 엑셀파일로 만들어 저장, 새 파일로 오픈
5) [사람이 개입] 사람이 합격자들 중 지원동기를 파악해 최종 합격여부를 판단한 후, 별도의 셀에 최종합격 여부 기록, 전체셀 선택하여 복사
6) 2nd script 실행
"""

import datetime, openpyxl, os, re, requests, xlrd, numpy as np, pandas as pd
from bs4 import BeautifulSoup as bs
from os import listdir
from time import sleep

pd.set_option("mode.chained_assignment",  None) # <==== 판다스 경고를 끈다

passed_count_set = 30 #합격인원 수
waited_count_set = 10 #대기인원 수

cha_name = ""

def get_app_data():
    login_id = input("취창업캠프사이트 관리자 아이디를 입력해 주세요 : ")
    login_pw = input("취창업캠프사이트 관리자 패스워드를 입력해 주세요 : ")
    # cha_name = ""

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
            get_app_data()
        
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
            get_app_data()

        file_name_list = []

        #다운로드1 : 올해 신청자 전체엑셀 다운로드
        this_year = datetime.datetime.now().year
        print(str(this_year) + "년 전체엑셀 다운로드 중...")
        try:
            app_data_in_this_year = s.get(base_url+"index.php?mod=lecture&act=xls_tot&actType=reqData&lgCode=5&year="+str(this_year))
            this_ext = app_data_in_this_year.headers["Content-Disposition"].split(".")[-1]
            this_file_name = "./total_app_data_" + str(this_year) + "."+this_ext
            file_name_list.append(this_file_name)
            open(this_file_name, "wb").write(app_data_in_this_year.content)
        except Exception as e:
            print(str(this_year) + "년 전체엑셀 다운로드 중 오류가 발생하였습니다.", e)

        #다운로드2 : 전년도 신청자 전체엑셀 다운로드
        last_year = this_year-1
        print(str(last_year) + "년 전체엑셀 다운로드 중...")
        try:
            app_data_in_last_year = s.get(base_url+"index.php?mod=lecture&act=xls_tot&actType=reqData&lgCode=5&year="+str(last_year))
            this_ext = app_data_in_last_year.headers["Content-Disposition"].split(".")[-1]
            this_file_name = "./total_app_data_" + str(last_year) + "."+this_ext
            file_name_list.append(this_file_name)
            open(this_file_name, "wb").write(app_data_in_last_year.content)
        except Exception as e:
            print(str(last_year) + "년 전체엑셀 다운로드 중 오류가 발생하였습니다.", e)

        #다운로드3 : 지원동기 엑셀 다운로드
        student_page_url = ""
        for strong in strongs:
            if "수 : " in strong.get_text() and "접" in strong.get_text()[:1]:
                href = strong.parent.attrs["href"]
                student_page_url=href[2:len(href)]
                break

        if student_page_url == "":
            print("차수명이 잘못되어 교육생 정보 다운로드에 실패하였습니다.")
            print("차수명을 정확히 입력한 후 다시 시도해 주세요.")
            exit()
            
        print("교육생 정보 다운로드 폴더 생성중...")
        makeDownloadDirectory(cha_name)
        print("교육생 정보 다운로드 폴더 생성완료 : " + cha_name)

        student_data = s.get(base_url+student_page_url)
        student_data_soup = bs(student_data.text, "html.parser")
        student_data_links = student_data_soup.select(".bt_sheet")
        count = 0
        for href in student_data_links:
            count = count+1
            print("교육생 정보 다운받는 중..({})".format(count))
            try:
                this_href = href.attrs["href"]
                this_data = s.get(base_url+this_href[2:])
                this_ext = this_data.headers["Content-Disposition"].split(".")[-1]
                this_file_name = cha_name + "/" + cha_name + "_" + href.text + "."+this_ext
                file_name_list.append(this_file_name)
                open(this_file_name, "wb").write(this_data.content)
                print("교육생 정보 다운로드 완료({})".format(count))
            except Exception as e:
                print("교육생 정보 다운로드에 실패하였습니다. ({})", e)

    return cha_name, file_name_list

def cal_semi_score(cha_name, file_name_list):
    print("가채점을 시작합니다...")

    # 다운로드받은 엑셀파일 분석하여 가합격자 체크
    df = pd.read_html(file_name_list[2], header=0)[0]

    df["학부졸업(예정)년월"] = df["학부졸업(예정)년월"].str.replace(pat=r"[^\w\s]", repl= r" ", regex=True)
    df["적합여부"] = 1.0 # 적합여부 판단 컬럼 생성 후 1로 초기화
    df["가점"] = 0.0    
    df["총점(가점포함)"] = 0.0 # 총점컬럼 생성 후 0으로 초기화
    df["지원동기"] = ""
    df["합격"] = ""

    # 학부졸업평점 숫자형 타입 체크
    print("--------------------------------------------------------------")
    print("학부졸업평점 오류 검사중...")
    for i in range(len(df["학부졸업평점"])):
        try:
            _ = float(df["학부졸업평점"].loc[i])
        except ValueError:
            df["학부졸업평점"].loc[i] = 0.0
            print(df["성명"].loc[i] + " : 학부졸업평점 점수 오류. 학점을 0점으로 처리함")

    df = df.astype({"학부졸업평점":"float"}) # 총점 계산을 위해 float형 변경
    print("학부졸업평점 오류 검사 완료")

    # 졸업예정년월 체크하여 적합여부 처리
    print("--------------------------------------------------------------")
    print("만 34세 이하, 졸업여부, 학점, 가산점 체크하여 점수 계산중...")
    for i in range(len(df["성명"])):
        df = df.astype({"학부졸업(예정)년월": "str", "생년월일": "str"})

        #졸업예정년월 전처리
        df_grad_ym = df["학부졸업(예정)년월"].loc[i].strip().replace(" ","").replace("-","").replace(".","")
        if len(df_grad_ym) == 5:
            df_grad_ym = df_grad_ym[:4] + "0" + df_grad_ym[4:]
        elif len(df_grad_ym) == 8:
            df_grad_ym = df_grad_ym[:4] + df_grad_ym[4:6]
        elif len(df_grad_ym) == 6:
            pass
        else:
            df_grad_ym = input("졸업예정년월 형식 오류를 발견하였습니다. {} : [{}]에 대해 YYYYMM 형식으로 지금 입력해 주세요 : ".format(df["성명"].loc[i], df_grad_ym))
            df_grad_ym = df_grad_ym.strip().replace(" ","").replace("-","").replace(".","")
        df_grad_ym = df_grad_ym[:4] + "-" + df_grad_ym[4:]
        df["학부졸업(예정)년월"].loc[i] = df_grad_ym
    
        # 6개월 이내 졸업예정 체크
        today = datetime.datetime.today()
        try:
            converted_day = datetime.datetime.strptime(df_grad_ym,"%Y-%m")
        except Exception as e:
            print("원본데이터의 졸업예정일자({})가 잘못 되었습니다.".format(df_grad_ym), e)
        day_gap = today - converted_day

        #졸업예정자인 사람만 적합
        if int(day_gap.days) >= -180 :
            df["적합여부"].loc[i] = 1 # 적합 대상자일 경우 1로 셋팅(-180(6개월)보다 같거나 큰 경우)
        else :
            df["적합여부"].loc[i] = 0 # 부접합 대상자일 경우 0으로 셋팅
            print(df["성명"].loc[i] + " : 졸업(예정)자가 아니어서 부적합 처리함")
        
        # 만 34세 이하만 적합
        this_year = datetime.datetime.now().year
        this_birth = df["생년월일"].loc[i].strip().replace(" ","").replace("-","").replace(".","")
        if len(this_birth) != 8:
            this_birth = input("생년월일 형식 오류를 발견하였습니다. {} : [{}]에 대해 YYYYMMDD 형식으로 지금 입력해 주세요 : ".format(df["성명"].loc[i], this_birth))
            this_birth = this_birth.strip().replace(" ","").replace("-","").replace(".","")

        this_birth_year = int(this_birth[:4])
        if this_birth_year < this_year - 34:
            df["적합여부"].loc[i] = 0
            print(df["성명"].loc[i] + " : 만 34세 이상으로 부적합 처리함")
        df["생년월일"].loc[i] = this_birth

        # 졸업여부 가산점
        if int(day_gap.days) >= 0 :
            score = df["총점(가점포함)"].loc[i].copy()
            df["총점(가점포함)"].loc[i] = score + 1 # 졸업자 +1점
        else :
            score = df["총점(가점포함)"].loc[i].copy()
            df["총점(가점포함)"].loc[i] = score + 0.5 # 졸업예정자 +0.5점
    
        # 학점 총점에 추가
        score_total = df["총점(가점포함)"].loc[i].copy() + df["학부졸업평점"].loc[i].copy()
        df["총점(가점포함)"].loc[i] = score_total
        
        # 가산점 추가
        str_etc = df["기타해당사항"].str.replace(" ", "").copy() # 공백 제거
        df["기타해당사항"]= str_etc
        
        score_total = df["총점(가점포함)"].loc[i].copy()
        if df["기타해당사항"].isnull :
            pass
        else :
            df["총점(가점포함)"].loc[i] = score_total + 0.5
            print(df["성명"].loc[i] + " : 가산점 처리함(" + str_etc + ")")
    
        # 적합여부 확인
        score_total = df["총점(가점포함)"].loc[i].copy()
        pass_yn = df["적합여부"].loc[i].copy()
        df["총점(가점포함)"].loc[i] = score_total *  pass_yn
    print("졸업여부, 학점, 가산점 체크하여 점수 계산완료")

    print("--------------------------------------------------------------")
    print("이전 신청/접수 내역 체크하여 추가 가산점 부여중...(접수일 기준 1년 이내 내역만 체크)")
    df_weighted = add_weight(df, cha_name, file_name_list)
    df_weighted_sorted = df_weighted.sort_values(by="총점(가점포함)", ascending=False) # 총점 순으로 정렬
    df_weighted_sorted = df_weighted_sorted.reset_index(drop=True)
    print("이전 신청/접수 내역 체크하여 추가 가산점 부여완료")

    #가합격, 가대기, 가불합격 여부 입력
    print("--------------------------------------------------------------")
    print("(가)합격, (가)대기, (가)불합격 여부 입력중...")
    for i in range(len(df_weighted_sorted["총점(가점포함)"])):
        if i<=passed_count_set:
            df_weighted_sorted["합격"].loc[i] = "(가)합격"
        elif i>passed_count_set and i<=passed_count_set+waited_count_set:
            df_weighted_sorted["합격"].loc[i] = "(가)대기"
        else:
            df_weighted_sorted["합격"].loc[i] = "(가)불합격"
    print("(가)합격, (가)대기, (가)불합격 여부 입력완료")


    print("--------------------------------------------------------------")
    print("가채점 결과에 지원동기 내용 입력중...")

    file_name = file_name_list[3]
    df_moti = pd.read_html(file_name, header=0)[0]
    df_moti=df_moti.fillna(0)
    df_moti["교육지원동기 및 의지"]=df_moti["교육지원동기 및 의지"].map(lambda x: re.sub('-=+,#/\?:^$＇▷“.@*\"※~&%ㆍ!』\\‘|\(\)\[\]\<\>`\'…》]',"",str(x)))

    i=0
    moti_data = []
    for i in range(len(df_moti)):
        this_data={}
        this_data["성명"]=df_moti.iloc[i]["성명"]
        this_data["지원동기"]=df_moti.iloc[i]["교육지원동기 및 의지"]
        this_data["접수일"]=df_moti.iloc[i]["접수일"].replace(".","-")
        moti_data.append(this_data)

    for i in range(len(df_weighted_sorted["성명"])):
        for moti in moti_data:
            if df_weighted_sorted["성명"].loc[i] == moti["성명"] and df_weighted_sorted["접수일"].loc[i] == moti["접수일"]:
                df_weighted_sorted["지원동기"].loc[i] = moti["지원동기"]
    print("가채점 결과에 지원동기 내용 입력완료")


    #가채점 결과 엑셀로 저장
    result_final_file_name = cha_name + "/" + cha_name + "_가채점 결과.xlsx"
    df_weighted_sorted.to_excel(result_final_file_name,
        sheet_name = "Sheet1", 
        # na_rep = 'NaN',
        float_format = "%.2f",
        header = True, 
        index = True, 
        index_label = "id", 
        startrow = 0, 
        startcol = 0, 
        freeze_panes = (1, 0)
    )
    print("--------------------------------------------------------------")
    print("가채점 완료 → 파일로 저장하였습니다. 파일명 : {}".format(result_final_file_name))

def add_weight(df_org, cha_name, file_name_list): #가채점 시 총점 최종 계산 전 가점을 반영함
    #올해, 전년도 전체엑셀 열어 병합
    df_total_this_year = pd.read_html(file_name_list[0], header=0)[0]
    df_total_last_year = pd.read_html(file_name_list[1], header=0)[0]
    df = pd.concat([df_total_this_year,df_total_last_year], ignore_index=True)
    df_filtered = df.copy()
    df_filtered["가점"] = 0.0

    for i in df.index:
        #취소자 제외
        if df["구분.1"].loc[i]=="취소자":
            df_filtered = df_filtered.drop([i])
            
        #합격자 & 블랙리스트 제외
        if df["합격여부"].loc[i]=="합격" or df["합격여부"].loc[i]=="블랙리스트":
            df_filtered = df_filtered.drop([i])

    df_filtered["접수일(int)"] = df_filtered["접수일"].str.replace(".", "").astype(int)
    df_filtered["접수일 1년전(int)"] = df_filtered["접수일"].str.replace(".", "").astype(int)
    pd.to_numeric(df_filtered["접수일(int)"], errors="coerce").fillna(0).astype(int)
    pd.to_numeric(df_filtered["접수일 1년전(int)"], errors="coerce").fillna(0).astype(int)
    df_filtered["접수일 1년전(int)"] = df_filtered["접수일 1년전(int)"] - 10000

    #df_filtered 에서 cha_name에 해당하는 사람만 반복, 접수일 기준 1년 내에 신청한 건이 있다면 가점에 기록
    for i in df_filtered.index:
        if df_filtered["교육명"].loc[i] == cha_name:
            for i2 in df_filtered.index:
                if df_filtered["교육명"].loc[i2] != cha_name and df_filtered["연락처"].loc[i2] == df_filtered["연락처"].loc[i] and df_filtered["접수일(int)"].loc[i2] >= df_filtered["접수일 1년전(int)"].loc[i]:
                    df_filtered["가점"].loc[i] = df_filtered["가점"].loc[i] + 0.3

    weighted_dic = []
    for i in df_filtered.index:
        this_weighted_dic = {}
        if df_filtered["가점"].loc[i] > 0.9:
            df_filtered["가점"].loc[i] = 0.9
        if df_filtered["가점"].loc[i] > 0:
            this_weighted_dic["phone"] = df_filtered["연락처"].loc[i]
            this_weighted_dic["weight"] = df_filtered["가점"].loc[i]
            weighted_dic.append(this_weighted_dic)
            print(df_filtered["수강생이름"].loc[i] + " : 가점 {}점 추가".format(df_filtered["가점"].loc[i]))

    #현재 명단에 가점에 해당되는 인원이 있는지 전화번호로 체크
    for i in df_org.index:
        for weight in weighted_dic:
            if df_org["휴대폰"].loc[i] == weight["phone"]:
                df_org["가점"].loc[i] = weight["weight"]
    
    return df_org

def makeDownloadDirectory(cha_name):
    try:
        if not(os.path.isdir("./"+cha_name)):
            os.makedirs(os.path.join("./"+cha_name))
    except OSError as e:
        if e.errno != errno.EEXIST:
            print(cha_name + " : 폴더 생성에 실패하였습니다.")
            raise

if __name__ == "__main__":
    cha_name = input("과정명을 입력하세요 : ").strip()
    cha_name, file_name_list = get_app_data()
    cal_semi_score(cha_name, file_name_list)

    print("==============================================================")
    print("1. 엑셀파일에서 지원동기를 확인한 후 '합격'셀에 최종 합격여부를 입력해 주세요(합격/대기/불합격으로 구분하여 입력).")
    print("2. 전체셀을 선택한 다음 복사(Ctrl+C)해 주세요.")
    print("3. 복사가 되었다면 아래 탭에 있는 코드를 실행해 주세요.")