import cv2
import numpy as np

#카메라 실행, 최초 배경 로드
cap = cv2.VideoCapture(0)
_, back = cap.read()

#캡처를 실시할 영역 지정
shot_area_width, shot_area_height = 200, 200
shot_area_check_ref = int((shot_area_width * shot_area_height) * 0.1)
cap_h, cap_w, _ = back.shape
shot_x, shot_y = int((cap_w - shot_area_width)/2), int((cap_h - shot_area_height)/2)

#배경차분을 위한 블러링
back_gray = cv2.cvtColor(back, cv2.COLOR_BGR2GRAY)
back_blur = cv2.GaussianBlur(back_gray, None, 1.0)
fback = back_blur.astype(np.float32)
while True:
    ret, frame = cap.read()
    if ret:
        #캡처가 되는 영역을 표시
        cv2.rectangle(frame, (shot_x, shot_y, shot_area_width, shot_area_height), (50, 255, 50), 3) 
        #이동평균 배경차분을 위해 새로운 영상을 배경영상에 0.01만큼 누적하여 반영
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY) #배경에 녹이기 위해 gray처리
        blur = cv2.GaussianBlur(gray, None, 1.0) #블러처리된 새로운 영상 정보
        cv2.accumulateWeighted(blur, fback, 0.01) #기존 배경 모델에 새 영상을 누적하여 반영
        back = fback.astype(np.uint8) #float으로 결과가 계산되므로 uint8로 변환
        diff = cv2.absdiff(blur, back) #새롭게 반영된 배경 모델 차분
        _, diff_threshold = cv2.threshold(diff, 30, 255, cv2.THRESH_BINARY)
        cnt, _, stats, _ = cv2.connectedComponentsWithStats(diff_threshold)
        for i in range(1, cnt):
            (x, y, w, h, area) = stats[i] #변화가 감지된 부분의 정보
            #초콜렛 객체가 shot_area에 모두 진입한 경우에만 rectangle 생성
            if x >= shot_x and (x + w) <= shot_x + shot_area_width and w <= shot_area_width and y >= shot_y and h <= shot_area_height and (y + h) <= shot_y + shot_area_height: 
                if area < shot_area_check_ref: #라벨링한 면적이 shot_area_check_ref보다 작으면 제외
                    continue
                cv2.rectangle(frame, (x, y, w, h), (0, 0, 255))
        cv2.imshow("org", frame)
        cv2.imshow("diff", diff)
        if cv2.waitKey(20) == 27:
            break
    else:
        break

cap.release()
cv2.destroyAllWindows()