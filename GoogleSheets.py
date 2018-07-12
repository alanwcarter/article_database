from __future__ import print_function
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import pandas as pd

# Setup the Sheets API
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
store = file.Storage('credential.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
    creds = tools.run_flow(flow, store)
service = build('sheets', 'v4', http=creds.authorize(Http()))

# Call the Sheets API
SPREADSHEET_ID = '1FuwuNVhAT06VnktAxGrj1pbZythmMiKCXAvJoFCgf5E'


method = input("Read (r), Write (w): ")

if method=="r":
    RANGE_NAME = 'Sheet1!A2:E'
    #result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID,
    #                                             range=RANGE_NAME).execute()
    
    result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    
    values = result.get('values', [])
    if not values:
        print('No data found.')
#    else:
#        print('First, Second:')
#        for row in values:
#            # Print columns A and E, which correspond to indices 0 and 4.
#            print('%s, %s' % (row[0], row[1]))
    else:
        full_list = pd.DataFrame(data=values, columns=['Title', 'Author', 'Year', 'Publisher', 'Comments'])
#        print(full_list)
        comment_search = input('What tags to look up?: ')
        print(full_list[full_list['Comments'].str.contains(comment_search)])
    
            
if method=="w":
    RANGE_NAME = 'Sheet1!A2:E'
    
    file_name = input("Enter file: ")
    with open(file_name, "r") as f: # Reads the requested RIS file
        content = f.read()
        #print(content)
        
        which_author=content.find('A1')
        if which_author>0:
            author_tag="A1"
        else:
            author_tag="AU"
        
        # Defines the title, author, year, and publisher of the article
        details = ["T1  - ", author_tag+"  - ", "Y1  - ", "PB  - "]            
        # Separates the RIS data by indents, may not work for all
        indents = [i for i,x in enumerate(content) if x=="\n"]
        #print(indents)
        risValues=[]

        # Grabs the vaule after each entry in details, and before the indent     
        a = 0
        for det in details:
            i = indents[a]
            while i<content.find(det):
                a = a + 1
                i = indents[a]
                
            risValues.append(content[content.find(det)+6:i])
    
    # Add comment to end of the list
    SPREADSHEET_ID = '1FuwuNVhAT06VnktAxGrj1pbZythmMiKCXAvJoFCgf5E'
    comments = input("Add comments about paper: ")
    risValues.append(comments)
    resources = {
            "majorDimension": "ROWS",
            "values": [risValues]
            }
    service.spreadsheets().values().append(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME,body=resources, valueInputOption="USER_ENTERED").execute()    
    
    
    
