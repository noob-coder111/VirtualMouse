import cv2
import autopy
import gesturemodule as g
import numpy as np
import time
import math
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import webbrowser
# import ctypes
import pyautogui



# ctypes.windll.user32.LockWorkStation()
wCam,hCam=640,480
frameR=100
count=0
value=0
smoothening=5
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
#volume.GetMute()
#volume.GetMasterVolumeLevel()
volRange=volume.GetVolumeRange()

minVol=volRange[0]
maxVol=volRange[1]

vol=0
volbar=400
volper=0

pTime=0
plocX,plocY=0,0
clocX,clocY=0,0

cap =cv2.VideoCapture(0)
cap.set(3,wCam)
cap.set(4,hCam)

detector=g.handDetector(maxHands=1)
wScr,hScr=autopy.screen.size()


while True:
    success, img=cap.read()
    img=cv2.flip(img,1)
    img=detector.findhands(img)
    lmList2=detector.findPosition2(img,draw=False)
    lmList,bbox=detector.findPosition(img)
    

    if len(lmList)!=0:
        x1,y1=lmList[8][1:]
        x2,y2=lmList[12][1:]

        fingers=detector.fingersUp()
        # print(fingers)

        if fingers[1]==1 and fingers[2]==0:
            cv2.rectangle(img,(frameR,frameR),(wCam-frameR, hCam-frameR),
                            (255,0,255),2)
            x3=np.interp(x1,(frameR,wCam-frameR),(0,wScr))
            y3=np.interp(y1,(frameR//10,hCam-150),(0,hScr))

            clocX=plocX+(x3-plocX)/smoothening
            clocY=plocY+(y3-plocY)/smoothening


            autopy.mouse.move(clocX,clocY)
            cv2.circle(img,(x1,y1),15,(255,0,255),cv2.FILLED)
            plocX,plocY=clocX,clocY
            count=0

        if fingers[2]==1 and fingers[1]==0 :
            pyautogui.click(button = 'right')


        if fingers[1]==1 and fingers[2]==1 and fingers[0]==0:
            length,img,lineInfo= detector.findDistance(8,12,img)
            if length <30:
                cv2.circle(img,(lineInfo[4],lineInfo[5]),15,(0,255,0),cv2.FILLED)
                if value==0:
                    value+=1
                    autopy.mouse.click()
            elif length >30:
                value=0
            
            count=0
        if fingers[1]==1 and fingers[2]==1 and fingers[0]==1:
            length,img,lineInfo= detector.findDistance(8,12,img)
            if length <20:
                cv2.circle(img,(lineInfo[4],lineInfo[5]),15,(0,255,0),cv2.FILLED)
                x3=np.interp(x1,(frameR,wCam-frameR),(0,wScr))
                y3=np.interp(y1,(frameR//10,hCam-150),(0,hScr))
                clocX=plocX+(x3-plocX)/smoothening
                clocY=plocY+(y3-plocY)/smoothening
                pyautogui.mouseDown(button='left')
                autopy.mouse.move(clocX,clocY)
                plocX,plocY=clocX,clocY
                
            elif length >20:
                pyautogui.mouseUp(button='left')          

        if fingers[0]==1 and fingers[1]==1 and fingers[4]==0 and fingers[2]==0:
            cv2.rectangle(img,(50,100),(85,400),(255,0,0),2)
            cv2.rectangle(img,(50,int(volbar)),(85,400),(255,0,0),cv2.FILLED)
            cv2.putText(img, f'{int(volper)} %', (40, 450), cv2.FONT_HERSHEY_PLAIN, 2, (255,0,0), 2)
            x1,y1=lmList2[4][1],lmList2[4][2]
            x2,y2=lmList2[8][1],lmList2[8][2]
            cx,cy=(x1+x2)//2,(y1+y2)//2

            cv2.circle(img,(x1,y1),15,(255,0,255),cv2.FILLED)
            cv2.circle(img,(x2,y2),15,(255,0,255),cv2.FILLED)
            cv2.line(img,(x1,y1),(x2,y2),(255,0,255),3)
            cv2.circle(img,(cx,cy),15,(255,0,255),cv2.FILLED)
            length=math.hypot(x2-x1,y2-y1)
            
            vol=np.interp(length,[50,260],[minVol,maxVol])
            volbar=np.interp(length,[50,260],[400,100])
            volper=np.interp(length,[50,300],[0,100])
            volume.SetMasterVolumeLevel(vol, None)
            count=0

            
        if fingers[4]==1 and fingers[1]!=1:
            if count==0:
                count+=1
                webbrowser.open('http://google.com',new=2)

        
        if fingers[4]==1 and fingers[1]==1:
            # pyautogui.keyDown("ctrl","winleft")
            # pyautogui.press("o")
            # pyautogui.keyUp("ctrl","winleft")
            pyautogui.hotkey("ctrl","c")

        if fingers[0]==1 and fingers[1]==1 and fingers[4]==1:
            pyautogui.hotkey("ctrl","v")
            time.sleep(1)

                
    
    cTime=time.time()
    fps=1/(cTime-pTime)
    pTime=cTime
    cv2.putText(img,str(int(fps)),(20,50),cv2.FONT_HERSHEY_PLAIN,3,(255,0,0),3)



    cv2.imshow("Image",img)
    cv2.waitKey(1)