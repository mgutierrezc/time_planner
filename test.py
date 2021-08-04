"""
Project time management script

Author: Marco Gutierrez
MIT License

Usage:
- Asks user for project information and keywords
- Returns hours spent in a month on the project in
excel spreadsheet
"""

from __future__ import print_function
from datetime import datetime, timedelta
import pandas as pd
import os.path
from dateutil.parser import isoparse
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import base64
import email

# if modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly', 'https://mail.google.com/']

class GoogleInfoGatherer:
    """
    Gathers information from chosen Google service. Requires credentials from 
    respective user and API
    """

    def __init__(self, credentials, token):
        self.credentials = credentials # name of json file with credentials
        self.token = token # name of json file with token

    def credential_reader(self):
        """
        Reads credentials for google api and returns them in required
        format

        Input: None
        Output: credentials in google readable format
        """

        # inputting credentials
        creds = None

        # checking if credentials previously read
        if os.path.exists(self.token):
            creds = Credentials.from_authorized_user_file(self.token, SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials, SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(self.token, 'w') as token:
                token.write(creds.to_json())

        return creds
    
    def initializing_service(self, service, version):
        """
        Initializes Google service and returns it ready for usage

        Input: service name (str), version (str)
        Output: Service build
        """
        creds = self.credential_reader()

        # initializing the calendar service
        return build(service, version, credentials=creds)

    def initial_final_dates(self, service_name):
        """
        Asks user to input his preferred initial and final dates

        Input: Initial and final search dates
        Output: Formatted initial and final search dates for Google search
        """

        # inputting initial date
        while True:
            min_time = input("Por favor, ingrese una fecha inicial con el siguiente formato (dd/mm/aaaa): ")
            
            try:
                min_time_formatted = datetime.strptime(min_time, '%d/%m/%Y').isoformat() + 'Z'
                if service_name=="gmail":
                    min_time_formatted = datetime.strptime(min_time, '%d/%m/%Y').strftime('%Y/%m/%d')
                
                min_time_type = type(min_time_formatted)
            except ValueError:
                print("No ha ingresado una fecha en el formato correcto. \nPor favor, vuelva a intentarlo")

            else:
                break
            

        # inputting final date
        while True:
            max_time = input("\nPor favor, ingrese una fecha final con el siguiente formato (dd/mm/aaaa)\n"+\
                        "o escriba 'ahora' para usar la fecha actual: ")

            try:
                if max_time == "ahora":
                    max_time_formatted = datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
                    if service_name=="gmail":
                        max_time_formatted = datetime.utcnow().strftime('%Y/%m/%d')
                else: 
                    max_time_formatted = datetime.strptime(max_time, '%d/%m/%Y').isoformat() + 'Z'
                    if service_name=="gmail":
                        max_time_formatted = datetime.strptime(max_time, '%d/%m/%Y').strftime('%Y/%m/%d')

                max_time_type = type(max_time_formatted)

            except ValueError:
                print("No ha ingresado una fecha en el formato correcto. \nPor favor, vuelva a intentarlo")
            
            else:
                break
                
        print(min_time_formatted, max_time_formatted)
        return min_time_formatted, max_time_formatted

    def get_projects_data(self):
        """
        
        """


class GmailGatherer(GoogleInfoGatherer):
    def __init__(self, credentials, token):
        super().__init__(credentials, token)
        self.last_searched_results = None # list that stores latest searched items
    
    def time_query(self, min_time=None, max_time=None):
        """
        Prepares a time based query for the mail search

        Input: initial and final times
        Output: query in Google format for Gmail API
        """

        # if min/max time not inputted, ask user
        if min_time == None or max_time==None:
            min_time, max_time = self.initial_final_dates("gmail")

        return f"before: {max_time} after: {min_time}"

    def search_emails(self, query):
        """
        Searches for all emails following a given query

        Input: query in Google format for Gmail API
        Output: search results
        """

        service = self.initializing_service('gmail', 'v1')

        mail_id_result = service.users().messages().list(userId="me", q=query).execute()
        number_results = mail_id_result["resultSizeEstimate"]

        mail_results = []
        if number_results > 0:
            message_ids = mail_id_result["messages"]
            
            for id in message_ids:
                message = service.users().messages().get(userId="me", id=id["id"], format="raw").execute()
                message_raw = base64.urlsafe_b64decode(message["raw"].encode("ASCII"))
                message_str = email.message_from_bytes(message_raw)

                content_types = message_str.get_content_maintype()
                print("content_types =", content_types)

                if content_types == "multipart":
                    parsed_message = message_str.get_payload()[0]
                    mail_result.append(parsed_message.get_payload())
                else:
                    mail_result.append(message_str.get_payload())

        else:
            print("no results")

        self.last_searched_results = mail_results
        return mail_results

    def search_emails_keywords(self):
        """
        Searches for emails with an specific keyword

        Input: User project name, user keywords per project
        Output: List with user events
        """

        email_dict = {}
        self.get_events() # searching all events





gaaa = GmailGatherer("credentials.json", "token.json")
query = gaaa.time_query()
print(gaaa.search_emails(query))

