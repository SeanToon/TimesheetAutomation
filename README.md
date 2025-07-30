# TimesheetAutomation

For this project I will be writing a python script and assigning a cron job to automate the creation of a new Timesheet spreadsheet, each time a new pay period starts.

First, I had to navigate to the google cloud console, and create a new project. I named my new project Timesheet Automation.

<img width="1349" height="318" alt="image" src="https://github.com/user-attachments/assets/4f267bff-7d39-47ca-9ebe-a7b928c64b6b" />

Next, I had to enable Google Drive API to interact with files and Google Sheets API Services to actually edit the files in my Google Drive.

<img width="460" height="649" alt="image" src="https://github.com/user-attachments/assets/cfa44543-6670-4434-827d-612edd7409fd" />

(or just search for "Google Drive API" and "Google Sheets API", select and click enable)

I had to create an Oauth client so that my script could open my spreadsheets. Creating a new client consists of creating a key and token so that the file has constant permission to the folder and files.

<img width="893" height="275" alt="image" src="https://github.com/user-attachments/assets/741ea4d6-6ec7-453c-b609-ffb57458c03e" />

Under "Application Type" select "Desktop app" and then name your client. A .json file will be created, and it will need to be used in the python script. 

<img width="499" height="251" alt="image" src="https://github.com/user-attachments/assets/68f10feb-bf86-468e-a849-05e31f5971b7" />

Next, I was able to start writing the script
I began by making sure that I had the required packages:

pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

from __future__ import print_function
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from calendar import monthrange
from datetime import datetime, timedelta
import os

# -- CONFIG --	

SCOPES = ['https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/spreadsheets']
TEMPLATE_FILE_ID ='file id'
DEST_FOLDER_ID = 'folder id'

# -- AUTHENTICATE --

def authenticate():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request()) 
        else:
            flow = InstalledAppFlow.from_client_secrets_file('key.json', SCOPES)
            creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
    return creds

creds = authenticate()
drive_service = build('drive', 'v3', credentials=creds)
sheets_service = build('sheets', 'v4' , credentials=creds)

# -- Get start & end of pay period (1 - 15 and 16 - end of month)  --

# Todays Date
today = datetime.today()
year = today.year
month = today.month
employees = {'Ethan', 'Marcos', 'Kari', 'Sean'}

#---------------------------------------------------------
# Employee emails for permissions
# employee_emails = {'name': '<employee email
#---------------------------------------------------------

# Pay Period

if today.day <= 15:
    start = datetime(year, month, 1)
    end = datetime(year, month, 15)

else: 
    last_day = monthrange(year, month)[1]
    start = datetime(year, month, 16)
    end = datetime(year, month, last_day)

for name in employees:
    sheet_name = f"{name} - Timesheet {start.strftime('%b %d')} - {end.strftime('%b %d')}"

# -- Copy template populate cells and rename it --
    
    name_str = f"Name: {name}"
    pay_period_str = f"{start.strftime('%b %d')} - {end.strftime('%b %d')}"
    current_date = start
    date_rows = []

    new_file = drive_service.files().copy(
        fileId=TEMPLATE_FILE_ID,
        body={
            'name': sheet_name,
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'parents': [DEST_FOLDER_ID]
        }
    ).execute()

# -- To add permissions --
#    drive_service.permissions().create(
#        fileId=new_file['id'],
#        body={
#            'type': 'user',
#            'role': 'writer',
#            'emailAddress': employee_emails[name]
#        }
#   ).execute()

    sheets_service.spreadsheets().values().update(
        spreadsheetId=new_file['id'],
        range='Sheet1!A1',
        valueInputOption='USER_ENTERED',
        body={'values': [[name]]}
    ).execute()

    sheets_service.spreadsheets().values().update(
        spreadsheetId=new_file['id'],
        range='Sheet1!A2',
        valueInputOption='USER_ENTERED',
        body={'values': [[pay_period_str]]}
    ).execute()
   
    while current_date <= end:
        if current_date.weekday() < 5:
            date_rows.append([current_date.strftime('%m-%d-%Y')])
        current_date += timedelta(days=1)
    
    sheets_service.spreadsheets().values().update(
        spreadsheetId=new_file['id'],
        range=f'Sheet1!A5:A{4 + len(date_rows)}',
        valueInputOption='USER_ENTERED',
        body={'values': date_rows}
    ).execute()

print(f" Created new Timesheet: {sheet_name}")
print(f" Sheet link : https://docs.google.com/spreadsheets/d/{new_file['id']}/edit")
