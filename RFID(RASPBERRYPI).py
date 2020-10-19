import json
import gspread
import time
from oauth2client.service_account import ServiceAccountCredentials
from luma.core.interface.serial import i2c
from luma.core.render import canvas
from luma.oled.device import ssd1306,ssd1325,ssd1331,sh1106
from PIL import ImageFont, ImageDraw
import RPi.GPIO as GPIO

     
serial = i2c(port=1, address=0x3C)
device = ssd1306(serial, rotate=0)
    #BODY
try: #Try connecting to Sheets API
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('IDPCred.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open('IDP Proj RFID').sheet1
    print("Internet access")
    internetAccess = True
except httplib2.ServerNotFoundError:
    print("No internet")
    internetAccess = False

def readValues():
    '''
    names = (0, 'Inventory')    count = (0, '')
            (1, 'COKE')                 (1, '5')
            (2, '100+')                 (2, '4')
            (3, 'PEPSI')                (3, '3')
            (4, 'TOMATO')               (4, '2')
            (5, 'PORK')                 (5, '1')

    inventory = {'Inventory': '', 'COKE': '5', '100+': '4', 'PEPSI': '3', 'TOMATO': '2', 'PORK': '1'}
    '''
    #goes through master inventory list and arranges in dictionary formatted above
    inventory = {}
    names = []
    for name in enumerate(sheet.col_values(1)):
        names.append(name[1]) #saves the product name in a list as a cache
        for count in enumerate(sheet.col_values(2)):
            if names[0] == count[0]:
                inventory[names[1]] = count[1]
    return inventory, names


def displayValues():
    with canvas(device) as draw:
        font = ImageFont.truetype("usr/share/fonts/truetype/freefont/FreeMono.ttf", size = 13)
        count = 0
        for names in inventory.keys():
            draw.text((0,count), str(names), font=font, fill = 255)
            draw.text((100,count),str(inventory[names]), font = font, fill = 255)
            count += 10
            
lastInventory = {}
inventory, names = readValues()
def checkValue(): #pings the inventory count 
    for count in enumerate(sheet.col_values(2)):
        inventory[names[count[0]]] = count[1] #update the inventory
    return inventory
        
while True:
    inventory = checkValue()
    print(inventory)
    print(inventory != lastInventory)
    if(inventory != lastInventory):
        print("not equal!")
        displayValues()
        lastInventory = inventory.copy()#dictionary copy
    time.sleep(1)

#if inventory != lastInventory():

    


##    inventory[names] = 
##print(sheet.col_values(1))
##print(sheet.col_values(2))
