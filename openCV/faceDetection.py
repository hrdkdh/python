import sys
from cv2 import cv2

model = "res10_300x300_ssd_iter_140000.caffemodel"
config = "deploy.prototxt"
net = cv2.dnn.readNet(model, config)

if net.empty():
    print("얼굴인식에 필요한 소스파일이 없습니다.")
    sys.exit()

capture = cv2.VideoCapture(0)
# capture.set(cv2.CAP_PROP_FRAME_WIDTH, 300)
# capture.set(cv2.CAP_PROP_FRAME_HEIGHT, 300)

while True:
    ret, frame = capture.read()
    blob = cv2.dnn.blobFromImage(frame, 1, (300, 300), (104, 177, 123))
    net.setInput(blob)
    out = net.forward()

    detect = out[0, 0, :, :]
    (h, w) = frame.shape[:2]

    #200개 중 가장 확률 높은 하나만 출력 (i=0)
    for i in range(detect.shape[0]):
        confidence = detect[i, 2]
        if confidence > 0.5:
            x1 = int(detect[i, 3] * w)
            y1 = int(detect[i, 4] * h)
            x2 = int(detect[i, 5] * w)
            y2 = int(detect[i, 6] * h)

            cv2.rectangle(frame, (x1, y1), (x2, y2), (0, 255, 0))

    cv2.imshow("camera", frame)
    if cv2.waitKey(1) > 0:
        break

capture.release()
cv2.destroyAllWindows()
