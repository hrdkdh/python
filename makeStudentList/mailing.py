import os
import clipboard
import pyautogui as pag
import win32com.client as win32
import pygetwindow as gw
import sys

from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

def sendEmail(downloaded_folder_name, cha_name):
    check = checkReady()
    if check is True:
        ep_id = input("EP 아이디를 입력하세요 : ") #EP ID
        ep_pw = input("EP 비밀번호를 입력하세요 : ") #EP PSWD
        #메일발송을 위한 정보
        mail_subject = "취업캠프 교육생 입금정보 등록을 요청드립니다." #메일 제목
        mail_content = "안녕하세요, 산업혁신교육그룹입니다.\n\n취업캠프 교육생 입금을 위한 정보를\n첨부와 같이 송부드리오니 반영 부탁드리겠습니다.\n\n감사합니다.\n\n※본 메일은 자동화된 코드에 의해 발송되었습니다.\n\n" #메일 내용
        mail_reciever = "hoan3532@poscohrd.com" #메일 수신자    
        if ep_id != "" and ep_pw != "":
            driver = initDriver()
            connectEpMail(driver, ep_id, ep_pw)
            openMailWindow(driver)
            attachFiles(driver, os.getcwd()+"\\"+downloaded_folder_name.replace("/", "\\")+"신분증 사본_"+cha_name+".zip")
            attachFiles(driver, os.getcwd()+"\\"+downloaded_folder_name.replace("/", "\\")+"통장 사본_"+cha_name+".zip")
            writeMailContents(driver, mail_reciever, mail_subject, mail_content)
            print("메일 작성을 완료하였습니다.")
        else:
            print("EP아이디와 비밀번호를 입력해 주십시오.")
            sendEmail()
    else:
        sleep(3)

def checkReady():
    success = True
    win2 = gw.getWindowsWithTitle("Internet Explorer")
    if len(win2)>0:
        print("오류 : 인터넷 익스플로어창이 열려있어 진행할 수 없습니다")
        print("열려있는 인터넷 익스플로어창을 모두 닫은 다음 재시도해 주세요")
        print("!!!EP에 로그인되어 있을 경우 반드시 로그아웃한 후 창을 닫아 주세요!!!")
        success = False
    return success

def initDriver():
    #드라이버 인스턴스 생성
    driver = webdriver.Ie("IEDriverServer")
    return driver

def connectEpMail(driver, my_id, my_pw):
    url = "http://swpsso.posco.net/idms/U61/jsp/login/login.jsp"
    driver.get(url)
    driver.find_element_by_xpath("//*[@id='username']").send_keys(my_id) # id 입력 창에 id 입력
    driver.find_element_by_xpath("//*[@id='password']").send_keys(my_pw) # pw 입력 창에 pw 입력
    driver.find_element_by_xpath("//*[@id='loginsubmit']").click() # 로그인 버튼 클릭
    driver.implicitly_wait(10)

    #기존 로그인 경고창 뜰 경우 alert accept 처리
    try:
        alert=driver.switch_to.alert
        alert.accept()
    except:
        print("기존 로그인 경고창 없음. 계속 진행.")

    waitTime = 10
    print("EP초기화가 완료될 때까지 "+str(waitTime)+"초간 대기중입니다...")
    for s in range(1, waitTime+1):
        print(f'{s}\r', end="")
        sleep(1)

    print("\n대기 완료")

    #클릭 에러가 생길 수 있어 여러번 시도함
    for i in range(len(driver.window_handles)):
        success_check = False
        driver.switch_to.window(driver.window_handles[i])
        this_page_title = driver.title
        if "EP(Enterprise Portal)" in this_page_title:
            for i in range(3):
                try:
                    print("EP 메일 아이콘 클릭 시도중...")
                    driver.find_element_by_xpath("//*[@id='533982']").click()
                except:
                    print("EP 메일 아이콘 클릭 예외발생... 재시도중")
                else:
                    print("EP 메일 아이콘 클릭 완료")
                    success_check = True
                    break
            break

    if success_check is False:
        print("진행 중 오류가 발생하였습니다. 재시도해 주십시오.")
        sys.exit()
        
def openMailWindow(driver):
    for i in range(-1, len(driver.window_handles)):
        success_check = False
        driver.switch_to.window(driver.window_handles[i])
        this_page_title = driver.title
        if "Mail" in this_page_title:
            for i in range(3):
                try:
                    print("메일쓰기창 오픈 클릭 시도중...")
                    driver.find_element_by_xpath("//*[@id='Lnb']/div[1]/a").click()
                except:
                    print("메일쓰기 클릭 예외발생... 재시도중")
                else:
                    print("메일쓰기 클릭 완료")
                    success_check = True
                    break
            break

    if success_check is False:
        print("진행 중 오류가 발생하였습니다. 재시도해 주십시오.")
        sys.exit()

def attachFiles(driver, attatch_file_name):
    for i in range(-1, len(driver.window_handles)):
        success_check = False
        driver.switch_to.window(driver.window_handles[i])
        this_page_title = driver.title
        if "메일쓰기" in this_page_title:
            for i in range(3):
                try:
                    print("첨부파일 추가버튼 클릭 시도중...")
                    driver.find_element_by_xpath("//*[@id='write_send_info']/table[2]/tbody/tr[2]/td/div/a[1]").click()
                except:
                    print("첨부파일 추가버튼 클릭 예외발생... 재시도중")
                else:
                    print("첨부파일 추가버튼 클릭 완료")
                    success_check = True
                    break
            break

    if success_check is False:
        print("진행 중 오류가 발생하였습니다. 재시도해 주십시오.")
        sys.exit()

    sleep(2)
    clipboard.copy(attatch_file_name)
    pag.press("enter")

    sleep(3)
    pag.keyDown("ctrlleft")
    pag.press("v")
    pag.keyUp("ctrlleft")
    pag.press("enter")
    print("첨부파일 추가 완료")

def writeMailContents(driver, mail_reciever_address, mail_subject, mail_content):
    sleep(10)
    driver.implicitly_wait(30)
    
    #순서대로 메일수신자 입력란, 메일제목 입력란 xPath
    xPaths = [
        ["메일수신자 입력", "//*[@id='token-input-send_to']", mail_reciever_address],
        ["메일제목 입력", "//*[@id='write_send_info']/table[2]/tbody/tr[1]/td/input", mail_subject],
        ["메일내용 입력", "//*[@id='dext_body']", mail_content]
    ]
    
    for i in range(-1, len(driver.window_handles)):
        driver.switch_to.window(driver.window_handles[i])
        this_page_title = driver.title
        if "메일쓰기" in this_page_title:
            for x in range(0,2):
                success_check = False
                for i in range(3):
                    try:
                        print(xPaths[x][0]+" 시도중...")
                        driver.find_element_by_xpath(xPaths[x][1]).send_keys(xPaths[x][2])
                        if x is 0:
                            sleep(3)
                            pag.press("enter")
                    except:
                        print(xPaths[x][0]+" 예외발생... 재시도중")
                    else:
                        print(xPaths[x][0]+" 완료")
                        success_check = True
                        break
                if success_check is False:
                    print("진행 중 오류가 발생하였습니다. 재시도해 주십시오.")
                    sys.exit()
            sleep(5)
            break

    print(xPaths[2][0]+" 시도중...")  
    pag.press("tab")
    driver.switch_to.active_element.send_keys(mail_content)

    print(xPaths[2][0]+" 완료")

    #메일 발송버튼 클릭
    print("메일 발송버튼 클릭 시도중...")
    sleep(2)
    driver.find_element_by_xpath("//*[@id=\"memo_content\"]/div[1]/ul[1]/li[1]/a").click()
    sleep(3)
    alert = driver.switch_to.alert
    alert.accept()
    print("파일 업로드를 위해 60초간 대기중...")
    sleep(60) #파일 업로드를 위해 대기
    print("메일 발송완료")