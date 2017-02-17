from __future__ import print_function

import web
import os
import cgi

import httplib2

import time

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

urls = (
    # "/(.*)", "index"
    "/", "index",
    "/students.html", "students",
    "/Test1.html", "Test1",
    "/home.html", "home",
    "/teachers.html", "teachers",
    "/images/(.*)", "images"

)
app2 = web.application(urls, globals(),True)

render = web.template.render('../templates/')  # ../ relativepath


# raise web.seeother('/students')
# greeting ="Hello anyone home"
# return render.index(greeting2=greeting2)
class images:
    def GET(self, name):
        ext = name.split(".")[-1]

        ctype = {
            "png": "image/png",
            "jpg": "image/jpg",
            "gif": "image/gif",
            "ico": "image/x-icon"
        }
        if name in os.listdir('../templates/images'):
            web.header("Content-Type", ctype[ext])
            return open('../templates/images/%s'%name,"rb").read()
        else:
            raise web.notfound("images not found")


class index(object):
    def GET(self):
        greeting = " "
        return render.index(greeting=greeting)


class teachers(object):
    def GET(self):
        greeting = " "
        return render.teachers(greeting=greeting)


class Test1(object):
    def GET(self):
        greeting = " "
        return render.Test1(greeting=greeting)

class home(object):
    def GET(self):
        greeting = " "
        return render.home(greeting=greeting)


class students(object):
    
    


    # If modifying these scopes, delete your previously saved credentials
    # at ~/.credentials/sheets.googleapis.com-python-quickstart.json
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    CLIENT_SECRET_FILE = 'client_secret.json'
    APPLICATION_NAME = 'Google Sheets API Python Quickstart'
    #    Make this a global variable?
    global columnRef
    columnRef = {0:'A',1:'B',2:'C',3:'D',4:'E',5:'F',6:'G',7:'H',8:'I',  #dict to ref google sheets cols and rows
                 9:'J',10:'K',11:'L',12:'M',13:'N',14:'O',15:'P',16:'Q',17:'R',
                 18:'S',19:'T',20:'U',21:'V',22:'W',23:'X',24:'Y',25:'Z'}

    def get_credentials(self):
        """Gets valid user credentials from storage.
    
        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.

        Returns:
            Credentials, the obtained credential.
        """
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir,
                                       'sheets.googleapis.com-python-quickstart.json')

        store = oauth2client.file.Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
            flow.user_agent = APPLICATION_NAME
            if flags:
                credentials = tools.run_flow(flow, store, flags)
            else: # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials



    def GET(self):
        greeting = " "

        return render.students(greeting=greeting)

    def POST(self):
        form = web.input(name="Nobody", greet=" ")
        greeting = " " (form.name, form.id, form.key)
        # raise web.seeother('/students')
        return render.students(greeting=greeting)

    def LOSE(self):
        print('hello')    
        
    def get_full_range( self,sheetInfo ):
        """Gets the full range of the grid
        Takes a dictionary value

        Returns: range in A1 notation string
        """

        row = sheetInfo.get('sheets')[0].get('properties').get('gridProperties').get('rowCount')
        column = sheetInfo.get('sheets')[0].get('properties').get('gridProperties').get('columnCount')
    #    print(column)
        name = sheetInfo.get('sheets')[0].get('properties').get('title')
        newName = str(str(name) + '!' + 'A1:' + str(columnRef.get(column - 1)) + str(row))
    #    print(newName)
        return newName



    def convrt_cell_range( self,sheetInfo,rowIndex,columnIndex ):
        """Converts the cell range for the spreadsheet to A1 notation
        Takes a dictionary value, a row index, and a column index

        Returns: range in A1 notation string
        """

        column = columnRef.get(columnIndex)
        name = sheetInfo.get('sheets')[0].get('properties').get('title')

        return name + '!' + column + str(rowIndex) + ':' + column + str(rowIndex)


    def is_enrolled( self,values,str ):
        """validates SID
        Takes a list of list and a string value of SID

        Returns: row value, -1 if not enrolled
        """

        inClass = False
        row = 0
        while inClass != True and row < len(values):
            #print(values[row], row)
            newVal = values[row]
            if newVal:
                if newVal[0] == str: #this is a list with string
                    inClass = True
            row = row + 1

        if inClass == True:
            return row
        else:
            return -1

    def check_date( self,values,str ): #NOTE: It may be possible to change pc date to log in, check for this loophole
        """validates date
        Takes a list of list and a string value of SID

        Returns: column value, -1 if not enrolled
        """

        isDate = False
        wasDate = False
        column = 0
        while isDate != True and column < len(values[0]):
            #print(values[row], row)
            newVal = values[0][column]
         #   print(newVal)
            if newVal == str: #this is a list with string
                isDate = True
               # if values[0][column + 1] = None or values[0][column + 1] != '':
               #     wasDate = True
            column = column + 1
        if isDate == True and wasDate == False:
            return column - 1
        else:
            return -1

    def check_key( self,values, str):

        isKey = False
        row = 0
    #    print(values)
    #    print(len(values))
        while isKey != True and row < len(values):
            #print(values[row], row)
            newVal = values[row]
            if newVal: 
              #  print(newVal[0])
              #  print(str)
              #  print(newVal[0] == str)
                if newVal[0] == str: #this is a list with string 
                    isKey = True
            row = row + 1
    
        return isKey

    def check_if_attended( self,values, row, column ):

        isAttended = False
        #print(values[row - 1][column])
        #print(values[column - 1])
        if len(values[row - 1]) > column:
            if values[row - 1][column] == "Check":
                isAttended = True
        return isAttended   

    def POST(self):
        """Shows basic usage of the Sheets API.

        Creates a Sheets API service object and prints the names and majors of
        students in a sample spreadsheet:
        https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
        """
        form = web.input(name="Nobody", greet=" ")
        credentials = self.get_credentials()
        http = credentials.authorize(httplib2.Http())
        discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                        'version=v4')
        service = discovery.build('sheets', 'v4', http=http,
                                  discoveryServiceUrl=discoveryUrl)

        
        spreadsheetId = '1kquOjTzLcH4HLnJvGUNjpmKRZEb3fWPMuHLy4ag3cZ0' #Change Sheet ID HERE----------------------------
        sheetInfo = service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
        fullRange = '' + self.get_full_range( sheetInfo )
        title = sheetInfo.get('sheets')[0].get('properties').get('title')
        rowCount = sheetInfo.get('sheets')[0].get('properties').get('gridProperties').get('rowCount')
        rangeSID = title + '!D1:D' + str(rowCount)
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheetId, range=rangeSID).execute()
        values = result.get('values', [])
        studentID = str(form.id) #Input for student ID here----------------------------------------------------------
        requestedKey = str(form.key)     #Input for key here-----------------------------------------------------------------
        print(values)
        rowID = self.is_enrolled( values, str(studentID) ) #Change ID HERE----------------------CAN IGNORE-----------------------
      #  print(rowID)
        valueToPrint = [
            [
            'Check' 
            ]
        ]
        body = {
          'values': valueToPrint
        }
        a = 0
        if rowID == -1:
        #    print('This number is not registered')
	        a = 1
        else:
            columnCount = sheetInfo.get('sheets')[0].get('properties').get('gridProperties').get('columnCount')
           # print(str(columnRef.get(columnCount)) + ' ' + str(columnCount))
            result2 = service.spreadsheets().values().get(
                spreadsheetId=spreadsheetId, range= str(title) + '!A1:' + str(columnRef.get(columnCount-1)) + '1').execute()
            values2 = result2.get('values', [])
            columnID = self.check_date( values2, time.strftime('%m/%d/%y') )
            if columnID == -1:
            #   print('Incorrect Date')
                a = 2
            else:
                #key = '95816' #CHANGE KEY HERE--------------------------CAN IGNORE---------------------------------------------
           #     print(columnID)
           #     print(columnRef.get(columnID))
                result3 = service.spreadsheets().values().get(
                    spreadsheetId=spreadsheetId, range= str(title) + '!' + str(columnRef.get(columnID)) + '1:' + 
                    str(columnRef.get(columnID)) + str(rowCount)).execute()
                values3 = result3.get('values', [])
                keyCheck = self.check_key( values3, requestedKey )
                if keyCheck == False:
                #    print('Incorrect Key')
                    a = 3
                else:
                 #   print(rowID)
                    result4 = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=fullRange).execute()
                    values4 = result4.get( 'values' , [])
                 #   print(str(rowID) + ' ' + str(columnID))
                    attChecker = self.check_if_attended( values4, rowID, columnID)
                    if attChecker == True:
                    #    print('Already checked-in')
                        a = 4
                    else:
                        service.spreadsheets().values().update(spreadsheetId=spreadsheetId,range=self.convrt_cell_range( sheetInfo,rowID,columnID ),
                            valueInputOption='USER_ENTERED',body=body).execute()
        greeting = 'Signed In'
        print(str(a))
        if a == 1:
            greeting = 'Incorrect ID'
        elif a == 2:
            greeting = 'Incorrect Date'
        elif a == 3:
            greeting = 'Incorrect Key'
        elif a == 4:
            greeting = 'Already Signed In'
        else:
            a = 5
        form = web.input(name="Nobody", greet=" ")
        #greeting = "%s, %s,%s" % (form.name, form.id, form.key)
        # raise web.seeother('/students')
        #print(greeting)
        return render.students(greeting = greeting)

if __name__ == "__main__":
    app2.run()