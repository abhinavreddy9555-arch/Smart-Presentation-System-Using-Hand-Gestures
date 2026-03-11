import cv2
from cvzone.HandTrackingModule import HandDetector
import win32com.client
import time

# Camera setup
cap = cv2.VideoCapture(0)

# Hand detector
detector = HandDetector(detectionCon=0.7, maxHands=1)

# Open PowerPoint
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
powerpoint.Visible = True

presentation = powerpoint.Presentations.Open(
    r"C:\\path\\to\\presentation.pptx", WithWindow=True
)

presentation.SlideShowSettings.Run()

while True:

    success, img = cap.read()
    img = cv2.flip(img,1)

    hands, img = detector.findHands(img)

    if hands:
        hand = hands[0]
        fingers = detector.fingersUp(hand)

        # Previous slide
        if fingers == [1,0,0,0,0]:
            presentation.SlideShowWindow.View.Previous()
            time.sleep(1)

        # Next slide
        if fingers == [0,0,0,0,1]:
            presentation.SlideShowWindow.View.Next()
            time.sleep(1)

        # Close presentation
        if fingers == [1,1,0,0,1]:
            presentation.Close()
            break

    cv2.imshow("Camera", img)

    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()