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

# if modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

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

    def initial_final_dates(self):
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

                else: 
                    max_time_formatted = datetime.strptime(max_time, '%d/%m/%Y').isoformat() + 'Z'
                    max_time_type = type(max_time_formatted)

            except ValueError:
                print("No ha ingresado una fecha en el formato correcto. \nPor favor, vuelva a intentarlo")
            
            else:
                break
    
        return min_time_formatted, max_time_formatted


class CalendarGatherer(GoogleInfoGatherer):
    """
    Gathers and process information from Google Calendar
    """

    def __init__(self, credentials, token):
        super().__init__(credentials, token)
        self.last_searched_results = None # list that stores latest searched items


    def get_events(self, min_time=None, max_time=None):
        """
        Gets events in specified time frame

        Input: None
        Output: List of event objects during period
        """

        service = self.initializing_service('calendar', 'v3')

        # if min/max time not inputted, ask user
        if min_time == None or max_time==None:
            min_time, max_time = self.initial_final_dates()

        events_result = service.events().list(calendarId='primary', timeMin=min_time,
                                            timeMax=max_time, singleEvents=True,
                                            orderBy='startTime').execute()
        events = events_result.get('items', [])
        formatted_events = []

        if not events:
            print('No se encontraron eventos.')
        for event in events:
            
            current_event = {}
            
            # beginning of event
            start = event['start'].get('dateTime', event['start'].get('date'))
            current_event['start'] = start

            # end of event
            end = event['end'].get('dateTime', event['end'].get('date'))
            current_event['end'] = end

            # duration
            duration = isoparse(end) - isoparse(start)
            current_event['duration'] = duration 

            # summary
            current_event['summary'] = event['summary']

            # description (if any)
            if "description" in event:
                current_event['description'] = event['description']
            else:
                current_event['description'] = ""

            # print(start, end, event['summary'])
            
            formatted_events.append(current_event)

        self.last_searched_results = formatted_events # storing last search in attribute 
        return formatted_events

    # TODO: add special character cleaning for events and keywords
    def search_events_keywords(self):
        """
        Searches for events with an specific keyword

        Input: User project name, user keywords per project
        Output: List with user events
        """

        event_dict = {}
        self.get_events() # searching all events

        while True:
            project_name = input("\nIndique el nombre del proyecto a buscar o 'Buscar'"+\
                            "\nsi ya terminó de registrar todos sus proyectos: ")
            
            if project_name != "Buscar":
                # initializing project entry
                event_dict[project_name] = {}
                project = event_dict[project_name]

                # initializing project events
                project["events"] = []

                # initializing keywords
                keywords = input("\nIndique las palabras clave del proyecto en el"+\
                                "\nsiguiente formato 'palabra_1, palabra_2, ...': ")
                inputted_keywords = keywords.split(",")
                title_keywords = project_name.split(",")
                parsed_keywords = inputted_keywords + title_keywords

                # searching for inputted keywords in all events
                for event in self.last_searched_results:
                    # checking if event has keyword on summary or description
                    for keyword in parsed_keywords:
                        if keyword in event["summary"]:
                            project["events"].append(event)
                        
                        elif keyword in event["description"]:
                            project["events"].append(event)

            else:
                break

        return event_dict

    def project_time_exporter(self, search_results):
        """
        Exports gathered data from project events into an Excel spreadsheet.
        It'll be exported into an specific sheet with spent time per
        project information.

        Input: results from project search
        Output: None
        """

        # creating dict with event information for excel
        parsed_events = {}
        parsed_events["Project"] = []
        parsed_events["Time spent (hours:minutes)"] = []
        
        # checking time spent per project
        for project in search_results:
            if search_results[project]["events"]: # if events where found for project

                parsed_events["Project"].append(project) # creating project entry
                project_time_spent = timedelta(seconds=0)

                for event in search_results[project]["events"]: # calculating total time spent per project
                    project_time_spent += event["duration"]
                
                hours, reminder_minutes = divmod(project_time_spent.seconds, 3600)
                minutes = reminder_minutes/60

                parsed_time = str(hours)+":"+str(minutes)
                parsed_events["Time spent (hours:minutes)"].append(parsed_time)
            
            else:
                print(f"IMPORTANTE: No se han encontrado resultados para el proyecto {project}.")
                break

        # exporting results to excel
        results_db = pd.DataFrame.from_dict(parsed_events)
        return results_db
        #results_db.to_excel("Time_spent_on_projects.xlsx", sheet_name="Time_spent", index=False)


class GmailGatherer(GoogleInfoGatherer):
    def __init__(self, credentials, token):
        super().__init__(credentials, token)
        self.last_searched_results = None # list that stores latest searched items
    
    def search_emails(self, min_time=None, max_time=None):
        """
        Searches for all emails in a given time frame
        """

        service = self.initializing_service('gmail', 'v1')

        # if min/max time not inputted, ask user
        if min_time == None or max_time==None:
            min_time, max_time = self.initial_final_dates()

        mail_result = service.users().search_messages().execute()

# initializing calendar gatherer and exporting results
user_calendar = CalendarGatherer("credentials.json", "token.json")
user_calendar_events = user_calendar.search_events_keywords()
user_calendar.project_time_exporter(user_calendar_events)