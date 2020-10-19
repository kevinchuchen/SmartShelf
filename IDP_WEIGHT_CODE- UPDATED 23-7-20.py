import json
import time
import gspread
from datetime import datetime
import sys
from oauth2client.service_account import ServiceAccountCredentials
import httplib2
import urllib3
import RPi.GPIO as GPIO
from hx711 import HX711
from luma.core.interface.serial import i2c
from luma.core.render import canvas
from luma.oled.device import ssd1306, ssd1325, ssd1331, sh1106
from PIL import ImageFont, ImageDraw


serial = i2c(port=1, address=0x3C) #GPIO2 = SDA, GPIO3 = SCL
device = ssd1306(serial, rotate=0)


def getWeight():
    val = int(hx.get_weight(5)) - 14295 #OFFSET WEIGHT
    print ('old val: ' + str(val))
    if val < 0:
        val = 0
    quantity = round(val/ITEMWEIGHT)
    hx.power_down()
    hx.power_up()
    time.sleep(0.1)
    print(val)
    return val, quantity

def cleanAndExit():
    print("Cleaning...")
    GPIO.cleanup()
    sys.exit()

def writeHeader():
    if internetAccess == True: #internet access, do cloud and local upload
        # updates for app
        sheet.update_cell(1, 2, "UPDATES") #row, column
        sheet.update_cell(2, 1, str(datetime.now().strftime("%d-%m-%Y")))
        sheet.update_cell(2, 2, "Weight(g)")
        sheet.update_cell(2, 3, "Quantity")
        
        C_lastElementCol = len(sheet.row_values(1)) #column for Cloud
        C_newElementCol = C_lastElementCol+2
       
        #HEADER FOR EACH DAY (Cloud only)
        sheet.update_cell(1, C_lastElementCol+2, str(datetime.now().strftime("%d-%m-%Y"))) #row, column
        sheet.update_cell(1, C_lastElementCol+3, "Weight(g)") #row, column
        sheet.update_cell(1, C_lastElementCol+4, "Quantity") #row, column

    return C_newElementCol

def oledDisplay(val, itemCount):
    with canvas(device) as draw:
        font = ImageFont.truetype("/usr/share/fonts/truetype/freefont/FreeMono.ttf", size = 15)
        draw.text((0, 0), str(val) + ' g', font = font,fill=255)
        draw.text((0,15), str(itemCount) +' Coke(s)', font = font, fill = 255)
        font = ImageFont.truetype("/usr/share/fonts/truetype/freefont/FreeMono.ttf", size = 15)
        
        draw.rectangle((0,30,120,60), outline="white", fill="black")
        draw.text((20,30), "Warehouse", font = font, fill = 255)
        draw.text((15,43), "254 Coke(s)", font = font, fill = 255)


    
#BODY

ITEMWEIGHT = 1000 #item weight in grams
referenceUnit = 24
hx = HX711(14, 15)#DAT to GPIO 14, SCK to GPIO 15
hx.set_reference_unit(referenceUnit)
hx.reset()
#hx.tare() #TARE ONLY WHEN FINDING REFERENCE
    
#API Request
try:
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('IDPCred.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open('IDP Proj').sheet1
    print("Internet access detected")
    internetAccess = True
except httplib2.ServerNotFoundError:
    print("No internet")
    internetAccess = False

C_newElementCol = writeHeader()
oldQuantity = 0

while True:
    weight, newQuantity = getWeight()
    oledDisplay(weight, newQuantity)
    if oldQuantity != newQuantity:
        print('old = ' + str(oldQuantity))
        print('new = ' + str(newQuantity))
        C_lastElementRow = len(sheet.col_values(C_newElementCol)) + 1
        sheet.update_cell(C_lastElementRow, C_newElementCol, str(datetime.now().strftime("%H:%M:%S"))) #row, column for date input
        sheet.update_cell(C_lastElementRow, C_newElementCol+1, str(weight) + "g") #Weight Input
        sheet.update_cell(C_lastElementRow, C_newElementCol+2, str(newQuantity) + " items") #quantity Input
    
        #update first column
        sheet.update_cell(3, 1, str(datetime.now().strftime("%H:%M:%S"))) #row, column for date input
        sheet.update_cell(3, 2, str(weight) + "g") #Weight Input
        sheet.update_cell(3, 3, str(newQuantity) + " items") #quantity Input
        
        oldQuantity = newQuantity
        
cleanAndExit()

