import cv2
import numpy as np
import pyautogui
import pytesseract
from win32 import win32gui
import json
import time
import gspread
import sys
from datetime import datetime, date, timedelta
from workalendar.asia import Malaysia
from oauth2client.service_account import ServiceAccountCredentials
import httplib2
import urllib3
import os.path
from pytrends.request import TrendReq
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import os
import random
import requests
import ast

def weatherPrediction():
    url = "https://community-open-weather-map.p.rapidapi.com/forecast"
    querystring = {"q":"Cheras"} #AREA TO MEASURE
    headers = {
        'x-rapidapi-host': "",
        'x-rapidapi-key': ""
        }

    response = requests.request("GET", url, headers=headers, params=querystring)

    print("place : " + str(querystring['q']))

    #Today's date
    datestr = '2020-08-15'
    obj = datetime.strptime(datestr, '%Y-%m-%d').date()
    todayDate = date.today() #obj


    #clear cells before updating
    sheetPred.update_cell(5,1,"- - --------------------------------------------------------------------------------------------- - -")
    sheetPred.update_cell(5,2,' ')
    sheetPred.update_cell(5,3,' ')
    cellRange = sheetPred.range('A6:C14')
    for cell in cellRange:
        cell.value = ""
        
    sheetPred.update_cells(cellRange)

    #Write Header
    sheetPred.update_cell(6,1, str(todayDate))
    sheetPred.update_cell(6,2, "Weather")
    sheetPred.update_cell(6,3, "Prediction")

    #prediction items
    prediction = {"clear sky": "Cold Drinks",
                  "scattered clouds": "Umbrella",
                  "overcast clouds": "Warm Drinks",
                  "broken clouds": "Umbrella",
                  "light rain": "Tissue Paper",           
                  "moderate rain":"Umbrella",
                  "heavy intensity rain": "Raincoat"}

    count = 7
    for entries in range(40):
        #Get the day and time in str format and convert it to ISO standard for comparison
        dayTime = ast.literal_eval(response.text).get('list')[entries].get('dt_txt')
        dayTime = datetime.strptime(dayTime, '%Y-%m-%d %H:%M:%S') 
        
    ##    if datetime.now() < dayTime: #remove past day and time
        #Get the day and time
        PredDate = datetime.strptime(ast.literal_eval(response.text).get('list')[entries].get('dt_txt').split(' ')[0], '%Y-%m-%d').date()
        PredTime = datetime.strptime(ast.literal_eval(response.text).get('list')[entries].get('dt_txt').split(' ')[1], '%H:%M:%S').time()

        #clear sky, scattered clouds, overcast clouds, broken clouds, light rain, moderate rain, heavy intensity rain
        description = ast.literal_eval(response.text).get('list')[entries].get('weather')[0].get('description')
        if todayDate == PredDate:
            sheetPred.update_cell(count, 1, str(PredTime))
            sheetPred.update_cell(count, 2, description)
            sheetPred.update_cell(count, 3, "Stock up on " + prediction[description])
            count+=1

def calendarPrediction():
    #Write header
    sheetPred.update_cell(1, 1, "Date")
    sheetPred.update_cell(1, 2, "Holiday")
    sheetPred.update_cell(1, 3, "Prediction")


    #Prediction items
    items = ['drinks', 'snacks', 'sanitizers', 'fruits', 'meat', 'seafood']


    holList = []
    #Next 3 coming holidays from today
    for holidays in Malaysia().holidays(datetime.now().year): #check all holidays in the current year
         #Neglecting past holidays and including the holiday itself
        if datetime.date(datetime.now()) < holidays[0] or datetime.date(datetime.now()) == holidays[0]:
            holList.append(holidays) #keeping holidays in a list
            
    i = 2
    if len(holList) > 3: #finding latest 3 holidays
        for hol in holList[0:3]:
            sheetPred.update_cell(i, 1, str(hol[0]))
            sheetPred.update_cell(i, 2, str(hol[1]))
            sheetPred.update_cell(i, 3, 'Stock up on %s' % (random.choice(items))) #random prediction
            i+=1
    else: #for 3 holidays or less
        for hol in holList:
            sheetPred.update_cell(i, 1, str(hol[0]))
            sheetPred.update_cell(i, 2, str(hol[1]))
            sheetPred.update_cell(i, 3, 'Stock up on %s' % (random.choice(items))) #random prediction        
            i+=1    


def googleTrends():
    pictureName = 'trendLine.png'
    ######### search keywords and trends #########
    pytrends = TrendReq(hl='en-MY')
    #Keyword to compare
    kw_list = ["Thermometer",'Sanitizer',"Face mask"]
    #configuration
    pytrends.build_payload(kw_list, cat=0, timeframe='today 12-m', geo='MY', gprop='')
    #get trendline
    data = pytrends.interest_over_time()
    #remove 'isPartial' column
    data = data.drop(labels=['isPartial'], axis='columns')
    image = data.plot(title= 'Trendline of Keywords')
    fig = image.get_figure()
    fig.savefig(pictureName)


    ######## Authentication and trendline upload ########
    g_login = GoogleAuth()
    #try to load credentials
    g_login.LoadCredentialsFile("mycreds.txt")
    if g_login.credentials is None:
        # Authenticate if they're not there
        g_login.LocalWebserverAuth()
    elif g_login.access_token_expired:
        #refresh if expired
        g_login.Refresh()
    else:
        #Initialize creds
        g_login.Authorize()
    g_login.SaveCredentialsFile("mycreds.txt")

    #removing old file and replacing with updated
    drive = GoogleDrive(g_login)
    file_list = drive.ListFile({'q': "'' in parents and trashed=False"}).GetList()
    for files in file_list:
        print(files['title'])
        if files['title'] == pictureName:
            files.Delete()
            trendLine = drive.CreateFile({'parents': [{'id': ''}]})
            trendLine.SetContentFile(pictureName)
            trendLine.Upload()
            break
        else:
            trendLine = drive.CreateFile({'parents': [{'id': ''}]})
            trendLine.SetContentFile(pictureName)
            trendLine.Upload()
            break
    print('Created file %s with mimeType %s' % (trendLine['title'], trendLine['mimeType']))


def screencap(window_title):
    if window_title:
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd:
            win32gui.SetForegroundWindow(hwnd)
            x, y, x1, y1 = win32gui.GetClientRect(hwnd) #get client size
            x, y = win32gui.ClientToScreen(hwnd, (x, y))
            x += 200
            y += 50 #cropping out top part
            x1, y1 = win32gui.ClientToScreen(hwnd, (x1 - x, y1 - y))
##            print(x, y, x1, y1)
            x1 -= 400 #cropping out right side
            y1 -= 580 #cropping out bottom part
            im = pyautogui.screenshot(region=(x, y, x1, y1))
            im.save(r'D:\Studies\IDP\SCREENSHOT.png')
        else:
            print('Window not found!')
    else:
        im = pyautogui.screenshot()
        return im
def processImg(): #Reads screenshot and does pre-processing before saving and sending to OCR
    imgProcess = cv2.imread('SCREENSHOT.png') #read saved image and process with cv2
    imgProcess = cv2.resize(imgProcess, (340,196))#resize image(double)
    imgProcess = cv2.cvtColor(imgProcess, cv2.COLOR_BGR2GRAY) #convert to grayscale
    cv2.imwrite(r'D:\Studies\IDP\SCREENSHOT.png', imgProcess)#save image

def get_text(image):
    pytesseract.pytesseract.tesseract_cmd = r'D:\Program Files\Tesseract-OCR\tesseract.exe'
    return pytesseract.image_to_string(cv2.imread(image))

def readRFID():
    screencap('RR9 Series HF Reader Demonstration Software(CS) V8.8') #take screenshot and saves it
    processImg() #Image pre-processing
    UID = get_text('SCREENSHOT.png').split()#image recognition to get UID in string format
    return UID

def inventoryCount(UID, lastUID, product):
    addDiff = list(set(UID)-set(lastUID)) #finding UID of item that is added
    removedDiff = list(set(lastUID)-set(UID))#finding UID of item that is removed
    inventory['REMOVED'] = removedDiff
    inventory['ADDED'] = addDiff

#list(programmedUID.items()) data sctructure = [('COKE', ['90936C9C500104E0', 'SCFE6C9C500104E0'])]
    for keypair in range(len(programmedUID)): #going through the dictionary of programmed items
        for value in list(programmedUID.items())[keypair][1]: #extracting the values of each key
            for removedUID in removedDiff: #go through each removed item and updates the product list
                if (removedUID == value):
                    product.remove(list(programmedUID.items())[keypair][0]) #remove product name
                    
            for addedUID in addDiff: #go through each added item and updates the product list
                if (addedUID == value):
                    product.append(list(programmedUID.items())[keypair][0])#add product name
    
        for keys in list(programmedUID.keys()): #search through every key and calculate total number of each product
            inventory[keys] = product.count(keys)
            
    lastUID = UID
    return inventory, lastUID

def updateValues():
    row = 1
    for products in list(programmedUID.keys()):
        row += 1
        sheet.update_cell(row,2, inventory[products])#Updating the product quantity

    C_lastElementCol = len(sheet.row_values(1)) #update time, weight, quantity index.
    C_lastElementRow = len(sheet.col_values(C_lastElementCol)) + 1
    sheet.update_cell(C_lastElementRow, C_lastElementCol-3, str(datetime.now().strftime("%H:%M:%S"))) #row, column

    for UID in inventory['REMOVED']:
        sheet.update_cell(C_lastElementRow, C_lastElementCol-2, str(UID)) #row, column
        
        for keypair in range(len(programmedUID)): #going through the dictionary of programmed items
            for value in list(programmedUID.items())[keypair][1]: #extracting the values of each key
                if value == UID:
                    sheet.update_cell(C_lastElementRow, C_lastElementCol-1, list(programmedUID.items())[keypair][0]) #Adding product name

        sheet.update_cell(C_lastElementRow, C_lastElementCol, "REMOVED") #Adding product status
        C_lastElementRow += 1
        
    for UID in inventory['ADDED']:
        sheet.update_cell(C_lastElementRow, C_lastElementCol-2, str(UID)) #row, column

        for keypair in range(len(programmedUID)): #going through the dictionary of programmed items
            for value in list(programmedUID.items())[keypair][1]: #extracting the values of each key
                if value == UID:
                    sheet.update_cell(C_lastElementRow, C_lastElementCol-1, list(programmedUID.items())[keypair][0]) #Adding product name            
        sheet.update_cell(C_lastElementRow, C_lastElementCol, "ADDED") #Adding product status
        C_lastElementRow += 1

     
        
def writeHeaders(programmedUID):
    #Adding total inventory headers for empty spreadsheet
    C_lastElementCol = len(sheet.row_values(1)) #column count for Cloud
    #holidays() #look ahead for holidays, if holiday found, write to sheet.
    try:
        if(sheet.row_values(1)[0] == 'Inventory'):
            #HEADER FOR EACH DAY (Cloud only)
            sheet.update_cell(1, C_lastElementCol+2, str(datetime.now().strftime("%d-%m-%Y"))) #row, column
            sheet.update_cell(1, C_lastElementCol+3, "Product (UID)") #row, column
            sheet.update_cell(1, C_lastElementCol+4, "Product Name") #row, column
            sheet.update_cell(1, C_lastElementCol+5, "Status") #row, column

    except IndexError: #if 'inventory' isn't found in A1 of spreadsheet
        row = 1
        sheet.update_cell(row,1, "Inventory")
        for products in list(programmedUID.keys()):
            row += 1
            sheet.update_cell(row,1, products)#Updating the product list header
                   
        #HEADER FOR EACH DAY (Cloud only)
        sheet.update_cell(1, C_lastElementCol+4, str(datetime.now().strftime("%d-%m-%Y"))) #row, column
        sheet.update_cell(1, C_lastElementCol+5, "Product (UID)") #row, column
        sheet.update_cell(1, C_lastElementCol+6, "Product Name") #row, column
        sheet.update_cell(1, C_lastElementCol+7, "Status") #row, column
        
        



########################################  BODY   #####################################################
try: #Try connecting to Sheets API
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('IDPCred.json', scope)
    client = gspread.authorize(creds)
    sheetPred = client.open('Calendar and Weather').sheet1
    sheet = client.open('IDP Proj RFID').sheet1
    print("Internet access")
    internetAccess = True
except httplib2.ServerNotFoundError:
    print("No internet")
    internetAccess = False
    
lastUID = []
inventory = {}
product = []
programmedUID = {
    'COKE' : ['90936C9C500104E0'],
    '100+' : ['65936C9C500104E0'],
    'PEPSI' : ['3B936C9C500104E0'],
    'WATER' : ['2DFE6C9C500104E0'],
    'SARSI':['SCFE6C9C500104E0']
}

if internetAccess == True: #internet access, do cloud and local upload
    #googleTrends()
    print("Trends done...")
    newSheet = writeHeaders(programmedUID)
    print("Headers written...")
    #calendarPrediction()
    print("calender prediction done...")
    #weatherPrediction()
    print("weather prediction done...")
    
    while True:
        UID = readRFID() #get UID sorted to be able to compare for changes
        UID.sort()
        inventory, lastUID = inventoryCount(UID, lastUID, product) #output is dict form, eg: {'PEPSI':3, 'COKE':2,'REMOVED':UID,'ADDED':UID}
        print('REMOVED = '+ str(inventory['REMOVED']))
        print('ADDED = '+ str(inventory['ADDED']))
        if(len(inventory['REMOVED']) != 0 or len(inventory['ADDED'])!= 0): #update only when there is a change
            updateValues()
        time.sleep(2)

    



