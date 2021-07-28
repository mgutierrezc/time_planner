from __future__ import print_function
from datetime import datetime
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

creds = None

if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

service = build('calendar', 'v3', credentials=creds)

# Call the Calendar API
now = datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
print(now)

print("Getting all events starting on DATE")

while True:
    min_time = input("Por favor, ingrese una fecha inicial con el siguiente formato (dd/mm/aaaa):")

    try:
        min_time_formatted = datetime.strptime(min_time, '%d/%m/%Y').isoformat() + 'Z'
        min_time_type = type(min_time_formatted)

    except ValueError:
        print("No ha ingresado una fecha en el formato correcto. \nPor favor, vuelva a intentarlo")

    else:
        break

events_result = service.events().list(calendarId='primary', timeMin=min_time_formatted,
                                    timeMax=now,
                                    maxResults=20, singleEvents=True,
                                    orderBy='startTime').execute()
events = events_result.get('items', [])

if not events:
        print('No upcoming events found.')
for event in events:
    start = event['start'].get('dateTime', event['start'].get('date'))
    print(start, event['summary'])