import os
import sys
import cv2
import urllib
import zipfile
import requests
import numpy as np
from time import sleep
from shutil import copyfile
from datetime import datetime
from bs4 import BeautifulSoup as bs

download_path = None
download_path_for_rename = None
pic_image_orginal_path = None
pic_image_resized_path = None
id_image_download_path = None
account_image_download_path = None
introduction_download_path = None
image_resize_size = [800, 600] #height, width / 이미지 비율은 4:3으로 고정

def downloadStudentImages(cha_name):
    login_id = input("취창업캠프사이트 관리자 아이디를 입력해 주세요 : ") #취창업캠프 관리자 ID
    login_pw = input("취창업캠프사이트 관리자 패스워드를 입력해 주세요 : ") #취창업캠프 관리자 PSWD    
    print("youth.posco.com 접속...")
    base_url = "http://youth.posco.com/posco/_owner/"
    login_url = base_url+"index.php?act=login"
    login_data = {
        "wd_id": login_id,
        "wd_pw": login_pw
    }
    with requests.Session() as s:
        login_req = s.post(login_url, data=login_data)
        if login_req.status_code != 200:
            print(login_req.status_code)
            print("관리자 화면 로그인에 실패하였습니다.")
            sleep(2)
            downloadStudentImages(cha_name)
        cha_list_data = s.get(base_url+"index.php?mod=lecture&act=main&cate=&sField=&sValue="+cha_name)
        soup = bs(cha_list_data.text, "html.parser")
        table = soup.select("table")
        try:
            strongs = table[1].select("strong")
        except:
            print("취창업캠프사이트 아이디/비번을 잘못 입력하였습니다.")
            sleep(2)
            downloadStudentImages(cha_name)
 
        student_page_url = ""
        for strong in strongs:
            if "격 : " in strong.get_text() and "합" in strong.get_text()[:1]:
                href = strong.parent.attrs["href"]
                student_page_url=href[2:len(href)]
                break

        if student_page_url == "":
            print("차수명이 잘못되어 이미지 다운로드에 실패하였습니다.")
            print("차수명을 정확히 입력한 후 다시 시도해 주세요.")
            sleep(2)
            downloadStudentImages(cha_name)

        print("이미지 다운로드 폴더 생성중...")
        makeDownloadDirectory(cha_name)
        print("이미지 다운로드 폴더 생성완료")  

        print("이미지를 다운받는 중... 기다려 주세요(1~3분 소요).")       

        student_list_data = s.get(base_url+student_page_url)
        image_soup = bs(student_list_data.text, "html.parser")
        image_links = image_soup.select("a")
        count_number = 0
        for href in image_links:
            if href.get_text() is not None and ("증명사진" in href.get_text() or "신분증" in href.get_text() or "통장" in href.get_text() or "자기소개서" in href.get_text()):
                this_href = href.attrs["href"]
                this_image = s.get(base_url+this_href[2:len(this_href)])
                try:
                    this_ext = this_image.headers["Content-Disposition"].split(".")[-1]
                except:
                    print("["+href.parent.parent.findAll("td")[3].find("a").text.replace("/", "_")+"]님의 파일 확장자가 없어 JPG로 대체되었습니다. 추후 확인하기 바랍니다.")
                    this_ext = "jpg"
                this_image_name = href.parent.parent.findAll("td")[3].find("a").text.replace("/", "_")+"_"+href.parent.parent.findAll("td")[4].find("strong").text
                if "증명사진" in href.get_text():
                    open("./"+pic_image_orginal_path+this_image_name+"."+this_ext, "wb").write(this_image.content)
                if "신분증" in href.get_text():
                    open("./"+id_image_download_path+this_image_name+"."+this_ext, "wb").write(this_image.content)
                if "통장" in href.get_text():
                    open("./"+account_image_download_path+this_image_name+"."+this_ext, "wb").write(this_image.content)
                if "자기소개서" in href.get_text():
                    open("./"+introduction_download_path+this_image_name+"."+this_ext, "wb").write(this_image.content)
                count_number += 1
                print(f'{count_number}건 다운로드중...\r', end=" └")
    makeZipFile(download_path, id_image_download_path, "신분증 사본_"+cha_name+".zip")
    makeZipFile(download_path, account_image_download_path, "통장 사본_"+cha_name+".zip")
    makeZipFile(download_path, introduction_download_path, "자기소개서 모음_"+cha_name+".zip")
    cropImages()
    downloaded_folder_name, pic_image_resized_path = changeDownloadFolderName()
    print("이미지 다운로드를 완료하였습니다.")

    return True, downloaded_folder_name, pic_image_resized_path

def makeDownloadDirectory(cha_name):
    global download_path, download_path_for_rename, pic_image_orginal_path, pic_image_resized_path, id_image_download_path, account_image_download_path, introduction_download_path, finally_folder_name
    
    #폴더 생성 및 PPT 생성을 위한 정보
    now_datetime = str(int(datetime.now().timestamp()))
    download_root_path = "results/downloaded_images/"
    download_path = download_root_path+now_datetime+"/"
    download_path_for_rename = download_root_path+cha_name+"_"+now_datetime+"/"
    pic_image_orginal_path = download_path+"pic_image_orginal/"
    pic_image_resized_path = download_path+"pic_image_resized/"
    id_image_download_path = download_path+"id_images/"
    account_image_download_path = download_path+"account_images/"
    introduction_download_path = download_path+"introduction_files/"
    finally_folder_name = download_path

    for dir_path in [download_root_path, download_path, pic_image_orginal_path, pic_image_resized_path, id_image_download_path, account_image_download_path, introduction_download_path]:
        try:
            if not(os.path.isdir("./"+dir_path)):
                os.makedirs(os.path.join("./"+dir_path))
        except OSError as e:
            if e.errno != errno.EEXIST:
                print(dir_path + " : 폴더 생성에 실패하였습니다.")
                raise

def makeZipFile(zip_file_path, org_file_path, zip_file_name): #압축된 파일이 저장될 폴더 / 압축할 파일이 있는 폴더 / 압축된 파일명
    this_zip = zipfile.ZipFile(zip_file_path+zip_file_name, "w")
    for folder, subfolders, files in os.walk(org_file_path):
        for f in files:
            this_zip.write(os.path.join(folder, f), os.path.relpath(os.path.join(folder,f), org_file_path), compress_type = zipfile.ZIP_DEFLATED)
    this_zip.close()

def cropImages():
    print("이미지를 4:3 비율로 자르고 다듬는 중...")
    try:
        os.remove("./"+download_path+"temp")
    except:
        pass
    files = os.listdir("./"+pic_image_orginal_path)
    for f in files:
        if len(f.split("."))<2: #폴더일 경우는 처리하지 않음
            continue

        dst = "./"+pic_image_orginal_path+"temp"
        copyfile("./"+pic_image_orginal_path+f, dst)
        src = cv2.imread(dst, cv2.IMREAD_COLOR)
        if src is not None:
            resized = src
            resize_ref = ""
            #리사이즈
            if src.shape[0] > image_resize_size[0]:
                if src.shape[1] > image_resize_size[1]: #높이도 기준보다 크고, 너비도 기준보다 큰 경우  → 높이 기준으로 리사이즈
                    resize_ref = "width"
                else: #높이는 기준보다 크지만, 너비는 기준보다 작은 경우 → 높이 기준으로 리사이즈
                    resize_ref = "height"
            else:
                if src.shape[1] > image_resize_size[1]: #높이는 기준보다 적고, 너비는 기준보다 큰 경우 → 너비 기준으로 리사이즈
                    resize_ref = "width"
                else: #높이가 기준보다 적고, 너비도 기준보다 작은 경우 → 리사이즈 불필요
                    pass

            if resize_ref == "width":
                resized=cv2.resize(src, dsize=(int(image_resize_size[1]), int(src.shape[0]*(image_resize_size[1]/src.shape[1]))), interpolation=cv2.INTER_AREA)
            elif resize_ref == "height":
                resized=cv2.resize(src, dsize=(int(src.shape[1]*(image_resize_size[0]/src.shape[0])), int(image_resize_size[0])), interpolation=cv2.INTER_AREA)

            #resized = faceRecognition(resized)

            #먼저 너비를 기준으로 조정 너비=4, 높이를 3에 맞춤
            width = resized.shape[1]
            height = int((width*4)/3)
            crop_axis = "height"

            #조정해야 할 높이가 현재 높이보다 크다면 높이를 기준으로 조정함
            if height > resized.shape[0]: 
                height = resized.shape[0]
                width = int((height*3)/4)
                crop_axis = "width"
            
            #잘라내야 할 높이 혹은 너비는?
            if crop_axis == "height":
                crop_size = resized.shape[0] - height
                resized = resized[0:height,:] #아래만 자른다
            elif crop_axis == "width":
                crop_size = int((resized.shape[1] - width)/2)
                crop_size_left = crop_size
                crop_size_right = resized.shape[1]-crop_size
                if (resized.shape[1] - width) % 2 > 0:
                    crop_size_left = crop_size + 1
                resized = resized[:,crop_size_left:crop_size_right] #좌우를 동등하게 자른다

            # print("최종 너비 : {}".format(width)+", 최종 높이 : {}".format(height)+", 결과 너비 : {}".format(resized.shape[1])+", 결과 높이 : {}".format(resized.shape[0])+", 잘라내야 할 축 : "+crop_axis+", 잘라내야 할 사이즈 : "+str(crop_size))

            cv2.imwrite(pic_image_resized_path+"temp_resized"+"."+f.split(".")[1], resized)
            os.rename("./"+pic_image_resized_path+"temp_resized"+"."+f.split(".")[1], "./"+pic_image_resized_path+f.split(".")[0]+"_"+str(resized.shape[1])+"x"+str(resized.shape[0])+"."+f.split(".")[1])
        os.remove("./"+pic_image_orginal_path+"temp")

def changeDownloadFolderName():
    downloaded_folder_name = download_path_for_rename
    pic_image_resized_path = download_path_for_rename+"pic_image_resized/"
    try:
        os.rename(download_path, download_path_for_rename)
    except:
        downloaded_folder_name = download_path
        pic_image_resized_path = download_path+"pic_image_resized/"
    return downloaded_folder_name, pic_image_resized_path