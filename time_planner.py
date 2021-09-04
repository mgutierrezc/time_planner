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
from calendar import month_name
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
                min_time_formatted = datetime.strptime(min_time, '%d/%m/%Y').isoformat() + "-05:00"
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
                    max_time_formatted = datetime.utcnow().isoformat() + "-05:00" # 'Z' indicates UTC time

                else: 
                    max_time_formatted = datetime.strptime(max_time, '%d/%m/%Y').isoformat() + "-05:00"
                    max_time_type = type(max_time_formatted)

            except ValueError:
                print("No ha ingresado una fecha en el formato correcto. \nPor favor, vuelva a intentarlo")
            
            else:
                break
    
        self.initial_form_date = min_time_formatted
        self.final_form_date = max_time_formatted
        
        return min_time_formatted, max_time_formatted


class CalendarGatherer(GoogleInfoGatherer):
    """
    Gathers and process information from Google Calendar
    """

    def __init__(self, credentials, token):
        super().__init__(credentials, token)
        self.last_searched_results = None # list that stores latest searched items
        self.initial_form_date = None # formatted initial date
        self.final_form_date = None # formatted initial date

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

            # summary (title)
            current_event['summary'] = event['summary']

            # description (if any)
            if "description" in event:
                current_event['description'] = event['description']
            else:
                current_event['description'] = ""

            formatted_events.append(current_event)

        self.last_searched_results = formatted_events # storing last search in attribute 
        return formatted_events

    # TODO: add special character cleaning for events and keywords
    def search_events_keywords(self):
        """
        Searches for events with an specific keyword

        Input: User project name, user keywords per project
        Output: Dict with user events {'project_title': [event_1, event_2, ...]}
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
                        if (keyword in event["summary"] or keyword in event["description"])\
                            and event not in project["events"]:
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
        list_of_projects = []

        # appending projects
        for project in search_results:
            list_of_projects.append(project) # creating project entry

        # initializing dataframe for storing info
        projects_db = pd.DataFrame(index=list_of_projects)

        # creating entries per day
        search_length = isoparse(self.final_form_date) - isoparse(self.initial_form_date)
        for day in range(search_length.days + 1):
            date = isoparse(self.initial_form_date) + timedelta(days=day)
            string_date = date.strftime("%d-%m")

            projects_db[string_date] = timedelta(0) # column with zeros per date
            
        # checking time spent per project
        for project in search_results:
            # current_row = projects_db.loc[projects_db["Project"] == project]
            
            if search_results[project]["events"]: # if events where found for project                                
                for event in search_results[project]["events"]:                    
                    # calculating total time spent per project on a specific day
                    initial_date = isoparse(event["start"]).strftime("%d-%m")
                    projects_db.at[project, initial_date] += event["duration"] 
        
        # formatting time spent on project as hours           
        projects_db = projects_db.applymap(lambda cell: divmod(cell.seconds, 3600)[0])
        
        
        # changing current day to weekday and time labor format
        current_columns = list(projects_db.columns)
        
        # concatenating year
        current_year = str(datetime.today().year)
        current_columns = [day+f"-{current_year}" for day in current_columns] 

        # checking if date is weekday
        evaluated_dates = [datetime.strptime(day, "%d-%m-%Y") for day in current_columns]
        evaluated_dates = [[day, day.weekday()] for day in evaluated_dates]

        # changing day to weekend when that's the case
        evaluated_dates = [day[0].strftime("%d-%m") if day[1]<5 else "weekend" for day in evaluated_dates]

        # changing month from num to month name
        final_dates = []
        for date in evaluated_dates:
            date = date.split("-")
            if len(date) == 2:
                # get month
                month = date[1]
                month = month_name[int(month)][:3]
                
                final_dates.append(date[0] + "-" + month)
            else:
                final_dates.append(date[0])

        # updating columns
        projects_db.columns = final_dates

        # changing weekend values to empty
        projects_db["weekend"] = ""

        print("all output: ", projects_db)
        return projects_db

# initializing calendar gatherer and exporting results
user_calendar = CalendarGatherer("credentials.json", "token.json")
user_calendar_events = user_calendar.search_events_keywords()
results = user_calendar.project_time_exporter(user_calendar_events)
results["Total"] = results.sum(axis=1)
results.to_excel("Time_spent_on_projects.xlsx", sheet_name="Time_spent")