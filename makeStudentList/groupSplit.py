import math
import random
import pandas as pd
from time import sleep

group_member_number_limit = 6

def checkClipboard():
    print("엑셀 양식에 데이터를 입력한 다음 클립보드에 복사해 주세요.")
    go_on_sign = input("복사가 완료되면 엔터키를 눌러 주세요.")
    
    df = pd.read_clipboard()
    if len(df) < 1 or "성명" not in df.columns:
        print("잘못된 데이터를 복사했습니다.")
        sleep(2)
        checkClipboard()
    else:
        df["조"] = 0
        df["출력순서"] = 0
        df["숙소"] = 0
    return df

def splitGroup():
    df = checkClipboard()

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

    df.to_excel("result.xlsx", sheet_name="조편성 명단")
    print(df)
    
    # for order_cit in order_cit_arr:
    #     pivot = df.pivot_table(index="조", values="성명", columns=order_cit, aggfunc="count")
    #     print(order_cit)
    #     print(pivot)

if __name__ == "__main__":
    splitGroup()