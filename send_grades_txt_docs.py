from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import csv
import win32com.client as win32
import time

STUDENTS = []
SPREADSHEETS = {}
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

def populateSpreadsheetFromCSV(csvFileName, dictionary):
    with open(csvFileName) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            #ignore header line of csv
            if line_count != 0:
                dictionary[row[0]] = row[1]
            line_count += 1

def populateListFromCSV(csvFileName, listName):
    with open(csvFileName) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            #ignore header line of csv
            if line_count != 0:
                listName.append(row[0])
            line_count += 1

def createStudentList(classSheet):
    populateListFromCSV(classSheet, STUDENTS)

def createSpreadsheetDictionary():
    populateSpreadsheetFromCSV('google_sheet_ids.csv', SPREADSHEETS)

def getGoogleSheetData(student, googleSheetId, googleSheetTab):
    data = googleSheetTab + '!A2:C'

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    waitBetweenAPICalls = True
    while waitBetweenAPICalls:
        try:
            sheet = service.spreadsheets()
            waitBetweenAPICalls = False
        except:
            time.sleep(10)
    result = sheet.values().get(spreadsheetId=googleSheetId,
                                range=data).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
        return None
    else:
        for row in values:
            if not row:
                print('Student did not submit.')
                return None
            if student == row[0]:
                return {'mark': row[1], 'feedback': row[2]}

def createTxtDocument(tab, mark, weight, text, recipient):
    dir = os.path.join(os.getcwd(),tab)
    if not os.path.exists(dir):
        os.mkdir(dir)
    f = open(os.path.join(dir,recipient + ".txt"),"w+")
    f.write("Mark is " + mark + "/" + weight + "<br>\n")
    f.write(text)
    f.close()

def main():
    print("Choose a spreadsheet: ")
    createSpreadsheetDictionary()
    for x in SPREADSHEETS:
        print('\t'+x)
    classSheet = input()
    print("Enter a spreadsheet tab name:")
    googleSheetTab = input()
    createStudentList(classSheet + '.csv')
    print("Enter weight (default is /10)")
    weight = input()
    if weight == "":
        weight = "10"
    counter = 0
    for student in STUDENTS:
        print(student)
        data = getGoogleSheetData(student, SPREADSHEETS[classSheet], googleSheetTab)
        if data != None:
            createTxtDocument(googleSheetTab, data['mark'], weight, data['feedback'], str(counter) + "-" + student)
            counter += 1

main()