import sys
import mailing
import makeList
import groupSplit
import downloadImages
from time import sleep

download_complete = False
downloaded_folder_name = None
pic_image_resized_path = None
cha_name = input("처리하려는 차수명(시스템에 등록된 차수명)을 정확히 입력해 주세요 : ").strip() #다운받고자 하는 차수명(정확해야 함)

def selectFunc():
    global download_complete, downloaded_folder_name, pic_image_resized_path

    print(" ")
    print("차수명 : " + cha_name)
    print("===================================================================================================") 
    print("1 : 교육생 자동 조편성")
    print("2 : 교육생 이미지 일괄 다운로드(증명사진/신분증사본/통장사본)")
    print("3 : 교육생 입금정보 등록메일 자동발송(신분증사본/통장사본)")
    print("4 : 교육생 PPT명단 작성")
    print("5 : 교육생 명찰 제작")
    print("6 : 프로그램 종료")
    print("===================================================================================================") 
    func = input("사용할 기능의 번호를 입력한 후 엔터키를 눌러주세요.")
    if func not in ["1", "2", "3", "4", "5", "6"]:
        print("                                                              ")
        print("정확한 번호를 입력해 주세요.")
        print("                                                              ")
        sleep(2)
        selectFunc()
    else:
        if func == "1":
            print("교육생 자동 조편성을 실행합니다...")
            groupSplit.splitGroup()
            sleep(2)
            selectFunc()
        elif func == "2":
            if download_complete is True:
                check = input("이미 다운로드를 받았습니다. 다시 다운로드 받으시려면 Y를, 앞으로 돌아가려면 N을 입력한 후 엔터키를 눌러주세요.")
            if download_complete is False or (check is "Y" or check is "y"):
                print("교육생 이미지 일괄 다운로드를 실행합니다...(증명사진/신분증사본/통장사본)")
                download_complete, downloaded_folder_name, pic_image_resized_path = downloadImages.downloadStudentImages(cha_name)
                print("다음 폴더에 이미지를 다운로드 받았습니다 : "+downloaded_folder_name)
                sleep(2)
                selectFunc()
            else:
                selectFunc()
        elif func == "3":
            if download_complete is True:
                print("교육생 입금정보 등록메일 자동발송을 실행합니다...(신분증사본/통장사본)")
                mailing.sendEmail(downloaded_folder_name, cha_name)
                sleep(2)
                selectFunc()
            else:
                print("교육생 이미지를 먼저 다운로드 받아주세요. (2번 기능을 먼저 실행해 주세요)")
                sleep(2)
                selectFunc()
        elif func == "4":
            if download_complete is True:
                print("교육생 PPT명단 작성을 실행합니다...")
                makeList.makePPT(downloaded_folder_name, pic_image_resized_path, cha_name)
                sleep(2)
                selectFunc()
            else:
                print("교육생 이미지를 먼저 다운로드 받아주세요. (2번 기능을 먼저 실행해 주세요)")
                sleep(2)
                selectFunc()
        elif func == "5":
            print("교육생 명찰 제작을 실행합니다...")
            sleep(2)
            selectFunc()
        elif func == "6":
            print("프로그램을 종료합니다.")
            sys.exit()

if __name__ == "__main__":
    selectFunc()