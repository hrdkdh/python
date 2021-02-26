import numpy as np
import cv2
import os

dir = "C:/Users/POSCOUSER/Desktop/src/"
dst_dir = "C:/Users/POSCOUSER/Desktop/dst/"
file_list = os.listdir(dir)
complete_cnt = 0
error_cnt = 0
for file in file_list:
    img = cv2.imread(dir+file)
    h, w_org, _ = img.shape
    w_new = int((4*h)/3)
    padding_size = int((w_new - w_org)/2)
    height_add_arr = np.full((h, padding_size, 3), 255, np.uint8)
    dst = np.append(height_add_arr, img, axis=1)
    if (padding_size*2)+w_org != w_new:
        add_score = w_new - ((padding_size*2)+w_org)
        if add_score < 0:
            add_score = -add_score
        height_add_arr = np.full((h, padding_size+add_score, 3), 255, np.uint8)
    dst = np.append(dst, height_add_arr, axis=1)
    try:
        cv2.imwrite(dst_dir+file, dst)
        complete_cnt += 1
    except:
        error_cnt += 1

print("총 {}건 중 {}건 변환완료({}건 변환실패)".format(len(file_list), complete_cnt, error_cnt))