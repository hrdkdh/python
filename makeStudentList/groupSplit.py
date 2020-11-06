import random
import pandas as pd

group_member_number_limit = 6

def checkClipboard():
    print("===================================================================================================")    
    print("엑셀 양식에 데이터를 입력한 다음 클립보드에 복사해 주세요.")
    go_on_sign = input("복사가 완료되면 엔터키를 눌러 주세요.")
    
    df = pd.read_clipboard()
    if len(df) < 1 or "성명" not in df.columns:
        print("잘못된 데이터를 복사했습니다.")
        checkClipboard()
    else:
        df["조"] = 0
        df["출력순서"] = 0
        df["숙소"] = 0
    return df

def splitGroup(df):
    #총 조 갯수 계산
    total_group_cnt = round(len(df)/group_member_number_limit)

    group_member_count_dic = {}
    for i in range(total_group_cnt):
        group_member_count_dic[i+1] = 0

    #step1 : 무작위로 조배정
    for i in range(len(df)):
        while True:
            this_group = random.randint(1, total_group_cnt)
            if group_member_count_dic[this_group] < group_member_number_limit:
                df.loc[i, "조"] = this_group
                group_member_count_dic[this_group] += 1
                break

    print(df)
    print(group_member_count_dic)

    pivot_by_univ = pd.pivot_table(df, index = ["대학명"], values="조", aggfunc = "count").query("조 == 1")
    print(pivot_by_univ)
    exit()

    for i in range(total_group_cnt):
        pivot_by_univ.query("조 == '"+str(i)+"'")
        print(pivot_by_univ)

    # print(df)
    # print(group_member_count_dic)
    

if __name__ == "__main__":
    df = checkClipboard()
    splitGroup(df)