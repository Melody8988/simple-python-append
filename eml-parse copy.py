#!/usr/bin/python

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pprint
from apiclient import discovery
import email
import xlrd
import datetime, time
import os, os.path


pp = pprint.PrettyPrinter()


scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'client_secret.json', scope)
client = gspread.authorize(creds)

# TODO: verify correct path and file name
excelFilePath = '/Users/melodysmac/RawEmail/Last Months Wo Costs DeptID 10060 for OIT.wid_30_8050922874131727691.xlsx'
excelFileName = 'Last Months Wo Costs DeptID 10060 for OIT.wid_30_8050922874131727691.xlsx'

# TODO: update title 
sheet = client.open('OIT WO Costs to DeptID 10060 4-1-16 to 3-31-18.xlsx').sheet1 

# TODO: update spreadsheet url
spreadsheet_id = '1hUtiGoS26ZsQu92Y9sOYAA5TTs7MzemK73zKOCCknRA'

# TODO: update sheet name and range 
range_name = "Report 1!A2:Q"

def get_credentials():
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        'client_secret.json', scope)
    return creds



def get_stuff(rangeName, spreadsheetId):

    # set up the google stuff
    credentials = get_credentials()
    service = discovery.build('sheets', 'v4', credentials=credentials)

    # grab data - info will be in list called 'values'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    values1 = [[s.encode('ascii') for s in list]
               for list in values]  # gets us a list of lists non-unicode
    print (values1)
    return(values1)


def wrtgoogle(values, rangeName, spreadsheetId):

    credentials = get_credentials()
    service = discovery.build('sheets', 'v4', credentials=credentials)

    # appends, build in exception feedback if something goes wonky

    body = {"values": values}
    result = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id, range=range_name,
        valueInputOption="USER_ENTERED", body=body).execute()

    return ()


def main():

    data = file('Openworkorders.eml').read()
    msg = email.message_from_string(data)  # entire message
   
    if msg.is_multipart():
        for payload in msg.get_payload():
            bdy = payload.get_payload()
    else:
        bdy = msg.get_payload()

    attachment = msg.get_payload()[1]


    # open and save excel file to disk
    f = open('Last Months Wo Costs DeptID 10060 for OIT.wid_30_8050922874131727691.xlsx', 'wb')
    f.write(attachment.get_payload(decode=True))
    f.close()


    xls = xlrd.open_workbook(excelFilePath)

    for sheets in xls.sheets():
        list = []
        for rows in range(sheets.nrows):
            for col in range(sheets.ncols):
                list.append(str(sheets.cell(rows, col).value))

        splitData = split(list, 15)
        values = splitData

        print splitData

        # wrtgoogle(values, range_name, spreadsheet_id)
    
    if os.path.isfile(excelFilePath):
        os.remove(excelFileName)
    else:
        print("The file does not exist")


def split(arr, size):

    # splits huge arry into individual row arrays
    individualRowArrays = []
    while len(arr) > size:
        piece = arr[:size]
        individualRowArrays.append(piece)
        arr = arr[size:]

    individualRowArrays.append(arr)
    
    # delete irrelevent rows that are just labels or empty
    individualRowArrays.pop()   # removes last row
    individualRowArrays.pop(1)
    individualRowArrays.pop(0)
    
    for row in individualRowArrays:

        # column A returns empty in these spreadsheets, so pop first index out
        row.pop(0)

        # excel stores dates as serial dates, so need to format to mm/dd/yy
        serial = float(row[3])
        seconds = (serial - 25569) * 86400.0
        newDate = datetime.datetime.utcfromtimestamp(seconds)
        dateWithoutTimeZone = newDate.date()

        
        cleanDate = str(dateWithoutTimeZone)
       
        formattedDate = datetime.datetime.strptime(cleanDate, '%Y-%m-%d').strftime('%m/%d/%y')

        # formula to calculate for fiscal year
        month = datetime.datetime.strptime(cleanDate, '%Y-%m-%d').strftime('%m')
        year = datetime.datetime.strptime(cleanDate, '%Y-%m-%d').strftime('%Y')

        fiscalMonth = int(month) + 6

        if fiscalMonth > 12:
            year = int(year) + 1
        
        fiscalYear = int(str(year)[-2:])

        row.insert(3, fiscalYear)       # this will become FY column 
        row.insert(4, cleanDate)        # this will become Date column
        row[5] = formattedDate          # this will become WO Entry Date column
        row.append(row[15])


    return individualRowArrays


if __name__ == "__main__":
    main()
    