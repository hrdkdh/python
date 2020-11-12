import os
import pandas as pd
from time import sleep

group_member_number_limit = 6
filesave_root_path = None
filesave_path = None

def checkClipboard():
    print("===================================================================================================")     
    print("엑셀 양식에 교육생 정보 데이터를 입력한 다음 클립보드에 복사해 주세요.")
    print("[조편성 데이터 양식.xlsx] 파일을 참고하세요")
    _ = input("복사가 완료되면 엔터키를 눌러 주세요.")
    
    try:
        df = pd.read_clipboard()
        if df is None or len(df) < 1 or "성명" not in df.columns:
            print("잘못된 데이터를 복사했습니다.")
            sleep(2)
            checkClipboard()
        else:
            df["조"] = 0
            df["출력순서"] = 0
            df["숙소"] = 0
    except:
        print("데이터가 복사되지 않았습니다.")
        sleep(2)
        checkClipboard()        
    return df

def splitGroup(cha_name):
    if cha_name == "" or cha_name == None:
        print("차수명을 입력하지 않았습니다.")
        cha_name = input("차수명을 입력해 주십시오.")
        splitGroup(cha_name)

    print("차수명 : " + cha_name)
    df = checkClipboard()
    print("자동 조편성 후 엑셀파일로 작성하는 중...")
    print(df)
    #총 조 갯수 계산
    total_group_cnt = round(len(df)/group_member_number_limit)

    #우선순위대로 정렬
    order_cit_arr = ["대학명", "계열", "성별", "거주지_시", "거주지_도"]
    df.sort_values(by=order_cit_arr, inplace=True)
    df.reset_index(drop=True, inplace=True) #인덱스 리셋

    for i in range(group_member_number_limit):
        this_start_num = i*total_group_cnt
        this_end_num = this_start_num+total_group_cnt
        group_number = 0
        for idx in range(this_start_num, this_end_num):
            group_number += 1
            df._set_value(idx, "조", group_number)
    
    order_cit_arr2 = ["조", "성명"]
    df.sort_values(by=order_cit_arr2, inplace=True)
    df.reset_index(drop=True, inplace=True) #인덱스 리셋

    print_no = 0
    for idx in df.index:
        print_no = print_no+1
        if print_no > group_member_number_limit:
            print_no = 1
        df._set_value(idx, "출력순서", print_no)

    makeDownloadDirectory()

    df.to_excel(filesave_path+"조편성표_"+cha_name+".xlsx", sheet_name="조편성 명단")
    print("조편성 완료")
    print("다음 폴더에 조편성 파일을 저장하였습니다 : " + filesave_path+"조편성표_"+cha_name+".xlsx")
    
    # for order_cit in order_cit_arr:
    #     pivot = df.pivot_table(index="조", values="성명", columns=order_cit, aggfunc="count")
    #     print(order_cit)
    #     print(pivot)

def makeDownloadDirectory():
    global filesave_root_path, filesave_path

    #폴더 생성 및 PPT 생성을 위한 정보
    filesave_root_path = "results/"
    filesave_path = filesave_root_path+"조편성 결과/"

    for dir_path in [filesave_root_path, filesave_path]:
        try:
            if not(os.path.isdir("./"+dir_path)):
                os.makedirs(os.path.join("./"+dir_path))
        except OSError as e:
            if e.errno != errno.EEXIST:
                print(dir_path + " : 폴더 생성에 실패하였습니다.")
                raise