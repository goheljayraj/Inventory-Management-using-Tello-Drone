import math
import pygame
import cv2
import numpy as np
import pyzbar.pyzbar as pyzbar
from djitellopy import Tello
import xlwt
import xlrd
from xlwt import Workbook
from time import sleep

# Workbook is created
wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Rapidex Inverntory.xls')
# sheet1.write(0, 0, )
style = xlwt.easyxf('font: bold 1')
sheet1.write(0, 0, "Sr. no.", style)
sheet1.write(0, 1, "CODE", style)
sheet1.write(0, 2, "QTY", style)
wb.save('Rapidex Inventory.xls')

rb = xlrd.open_workbook('Rapidex Inventory.xls')
sheet = rb.sheet_by_index(0)
print(sheet.nrows)

# For row 0 and column 0
print(sheet.cell_value(0, 0))

index = 0
contained = False
iterator = 0

w, h = 960, 720

tello = Tello()

tello.connect()
print(tello.get_battery())
tello.streamoff()
tello.streamon()
# tello.takeoff()

# tello.move_left(20)
# sleep(2)
# tello.move_forward(50)

# tello.rotate_counter_clockwise(90)

# tello.land()


KNOWN_DISTANCE = 55.1  # inches
KNOWN_WIDTH = 5.0  # inches

# define the fonts
fonts = cv2.FONT_HERSHEY_COMPLEX
Pos = (50, 50)
# colors (BGR)
WHITE = (255, 255, 255)
BLACK = (0, 0, 0)
MAGENTA = (255, 0, 255)
GREEN = (0, 255, 0)
CYAN = (255, 255, 0)
GOLD = (0, 255, 215)
YELLOW = (0, 255, 255)
ORANGE = (0, 165, 230)

cap = cv2.VideoCapture(0)
cap.set(3, 640)
cap.set(4, 408)

pygame.init()
win = pygame.display.set_mode((200, 200))


running = True



def mission():

    unit = 'm'
    shelfHeight = 0.25
    shelfNumber = 4
    lenggth = 1

    if unit == 'ft':
        shelfHeight *= 0.3048
        lenggth *= 0.3048
        # conversion
    shelfHeight*=100
    shelfHeight = int(shelfHeight)
    if shelfNumber%2 == 0:
        while shelfNumber>=2:
            for x in range(0,lenggth):
                tello.move_right(50)
            tello.move_up(shelfHeight)
            for x in range(0,lenggth):
                tello.move_left(50)

            shelfNumber -= 2
            if shelfNumber != 0:
                tello.move_up(shelfHeight)
        tello.land()

    else:
        while(shelfNumber >= 1):
            if(shelfNumber == 1):
                for x in range(0,lenggth):
                    tello.move_right(50)
                shelfNumber -= 1
            else:
                for x in range(0, lenggth):
                    tello.move_right(50)
                tello.move_up(shelfHeight)
                for x in range(0, lenggth):
                    tello.move_left(50)
                shelfNumber -= 2
            if(shelfNumber != 0):
                tello.move_up(shelfHeight)

        for x in range(0,lenggth):
            tello.move_left(50)
        tello.land()






def getKey(keyName):
    ans = False
    for eve in pygame.event.get(): pass
    keyInput = pygame.key.get_pressed()
    myKey = getattr(pygame, 'K_{}'.format(keyName))
    if keyInput[myKey]:
        ans = True
        pygame.display.update()
        return ans


def getKeyboardInput():
    lr, fb, ud, yv = 0, 0, 0, 0
    speed = 30

    if getKey("LEFT"):
        lr = -speed
    elif getKey("RIGHT"):
        lr = speed

    if getKey("UP"):
        fb = speed
    elif getKey("DOWN"):
        fb = -speed

    if getKey("w"):
        ud = speed
    elif getKey("s"):
        ud = -speed

    if getKey("a"):
        yv = -speed
    elif getKey("d"):
        yv = speed

    if getKey("l"): tello.land()

    if getKey("t"): tello.takeoff()

    if getKey("m"): mission()

    return [lr, fb, ud, yv]


def eucaldainDistance(x, y, x1, y1):
    eucaldainDist = math.sqrt((x1 - x) * 2 + (y1 - y) * 2)

    return eucaldainDist


def distanceFinder(focalLength, knownWidth, widthInImage):
    distance = ((knownWidth * focalLength) / widthInImage)
    return distance


def focalLengthFinder(knowDistance, knownWidth, widthInImage):
    focalLength = ((widthInImage * knowDistance) / knownWidth)
    return focalLength


def DetectQRcode(image):
    codeWidth = 0
    x, y = 0, 0
    euclaDistance = 0
    global Pos
    # convert the color image to gray scale image
    Gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # create QR code object
    objectQRcode = pyzbar.decode(Gray)
    for obDecoded in objectQRcode:

        points = obDecoded.polygon
        if len(points) > 4:
            hull = cv2.convexHull(
                np.array([points for point in points], dtype=np.float32))
            hull = list(map(tuple, np.squeeze(hull)))
        else:
            hull = points

        n = len(hull)
        # draw the lines on the QR code
        for j in range(0, n):
            # print(j, "      ", (j + 1) % n, "    ", n)

            cv2.line(image, hull[j], hull[(j + 1) % n], ORANGE, 2)
        # finding width of QR code in the image
        x, x1 = hull[0][0], hull[1][0]
        y, y1 = hull[0][1], hull[1][1]

        Pos = hull[3]
        # using Eucaldain distance finder function to find the width
        euclaDistance = eucaldainDistance(x, y, x1, y1)

        # retruing the Eucaldain distance/ QR code width other words
        return euclaDistance


refernceImage = cv2.imread("referenceImage2.png")
Rwidth = DetectQRcode(refernceImage)

# finding the focal length
focalLength = focalLengthFinder(KNOWN_DISTANCE, KNOWN_WIDTH, Rwidth)
print("Focal length:  ", focalLengthFinder)

while True:
    vals = getKeyboardInput()
    tello.send_rc_control(vals[0], vals[1], vals[2], vals[3])
    myFrame = tello.get_frame_read().frame
    img = cv2.resize(myFrame, (w, h))
    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            running = False
        if running == False:
            pygame.quit()
    for barcode in pyzbar.decode(img):
        mydata = barcode.data.decode('utf-8')
        contained = False
        while iterator < sheet.nrows:
            if mydata == sheet.cell_value(iterator, 1):
                contained = True
                iterator = 0
                break
            iterator += 1
        iterator = 0
        if not contained:
            print(mydata)
            sheet1.write(index + 1, 0, index + 1)
            sheet1.write(index + 1, 1, mydata)
            sheet1.write(index + 1, 2, 1)
            index += 1
            wb.save('Rapidex Inventory.xls')
            rb = xlrd.open_workbook('Rapidex Inventory.xls')
            sheet = rb.sheet_by_index(0)
        # print(mydata)

        pts = np.array([barcode.polygon], np.int32)
        pts = pts.reshape((-1, 1, 2))
        # cv2.polylines(img, [pts], True, (255, 0, 255), 5)
        pts2 = barcode.rect
        cv2.putText(img, mydata, (pts2[0], pts2[1]), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (255, 0, 255), 2)
        codeWidth = DetectQRcode(img)
        if codeWidth is not None:
            # print("not none")
            Distance = distanceFinder(focalLength, KNOWN_WIDTH, codeWidth)
            # print(round(Distance, 2))
            cv2.putText(img, f"Distance: {round(Distance, 2)} Cm", (20, 20), fonts, 0.6, (CYAN), 2)
            # while Distance < 80:
            #     tello.send_rc_control(0, -30, 0, 0)
            #     myFrame = tello.get_frame_read().frame
            #     img = cv2.resize(myFrame, (w, h))
            #     codeWidth = DetectQRcode(img)
            #     if codeWidth is not None:
            #         # print("not none")
            #         Distance = distanceFinder(focalLength, KNOWN_WIDTH, codeWidth)
            #     print("backin up??????????????")

            # cv2.putText(img, f"Distnace: {round(Distance, 2)} CM", (pts2[0], pts2[1]), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (255, 0, 255), 2)
            # betterLook.showText(img, f"Distnace: {round(Distance, 2)} CM", Pos, GOLD, int(Distance * 4.5))
        # out.write(im)

    cv2.imshow('Result', img)
    cv2.waitKey(1)