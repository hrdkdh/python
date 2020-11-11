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

    group_dic_for_cal = []
    for i in range(total_group_cnt):
        group_dic_for_cal.append({"조" : i+1, "인원" : 0})

    #step1 : 무작위로 조배정
    for i in range(len(df)):
        while True:
            this_group = random.randint(1, total_group_cnt)
            breakCheck = False
            for group in group_dic_for_cal:
                if group["조"] ==  this_group and group["인원"] < group_member_number_limit:
                    df.loc[i, "조"] = this_group
                    group["인원"] += 1
                    breakCheck = True
                    break
            if breakCheck == True:
                break
    
    # age_average = round(df.mean(axis=0, skipna=True)["나이"])
    # age_averages_by_group = pd.pivot_table(df, index = ["조"], values="나이", aggfunc="mean")
    # print(age_average)
    # for i in range(len(age_averages_by_group)):
    #     print(round(age_averages_by_group.iloc[i]["나이"]))

    #step2 : 거주지역별로 균등배분
    count_by_area = pd.pivot_table(df, index = ["거주지(시)"], values="성명", aggfunc="count").to_dict() #거주지별 총인원
    group_count_by_area = pd.pivot_table(df, index = ["조"], values="성명", columns="거주지(시)", aggfunc="count") #조(행) 거주지(열) 매트릭스
    for key, _ in count_by_area["성명"].items(): #거주지별로 loop. key는 거주지, _는 거주지별 총 인원수(value)
        this_area_one_group = []
        this_area_over_group = []
        this_area_zero_group = []
        for key_by_group, value in group_count_by_area[key].items(): #각 거주지별로 조에 배당된 수를 체크. key_by_group는 조, value는 배당된 수
            if value == 1: #1로 배분된 조 배열에 저장
                this_area_one_group.append(key_by_group)
            if value > 1: #2 이상으로 배분(x)된 조 배열에 저장
                this_area_over_group.append(key_by_group)
            elif math.isnan(value): #0으로 배분된 조 배열에 저장
                this_area_zero_group.append(key_by_group)
        if len(this_area_over_group) > 0: #2 이상으로 배분(x)된 조가 있다면
            if len(this_area_zero_group) > 0: #0으로 배당된 조를 찾는다
                #해당 조로 이동하고, 이동한 만큼 해당 조의 1명을 현재 조로 옮겨온다
                pass
            else: #0으로 배당된 조가 없다면 x와 2차이 이상 나는 곳을 찾는다
                pass
        else: #x와 2차이 이상 나는 곳이 없다면 그대로 종료
            pass
        # print(this_area_one_group)
        # print(this_area_over_group)
        # print(this_area_zero_group)
        break
    
    
    
    print(count_by_area)
    print(group_count_by_area)
    exit()

    

    # writer = pd.ExcelWriter('df.xlsx', engine='xlsxwriter')
    # df.to_excel(writer, sheet_name='Sheet1')
    # writer.close()

    pivot_by_univ = pd.pivot_table(df, index = ["대학명"], values="조", aggfunc = "count").query("조 == 1")
    

    for i in range(total_group_cnt):
        pivot_by_univ.query("조 == '"+str(i)+"'")
        print(pivot_by_univ)

    # print(df)
    # print(group_dic_for_cal)

if __name__ == "__main__":
    splitGroup()