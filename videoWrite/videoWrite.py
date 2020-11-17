import cv2

cap = cv2.VideoCapture("running_kid_with_balloon.avi")
fps = cap.get(cv2.CAP_PROP_FPS)
w = round(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
h = round(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))

fourcc = cv2.VideoWriter_fourcc(*"H264")
out = cv2.VideoWriter("output.mp4", fourcc, fps, (w, h))

while True:
    ret, frame = cap.read()
    if ret:
        out.write(frame)
        cv2.imshow("frame", frame)
        if cv2.waitKey(int(fps)) == 27:
            break
    else:
        break

cap.release()
out.release()
cv2.destroyAllWindows()