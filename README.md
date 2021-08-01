# Time planner
**Author**: Marco Gutierrez

## Description
Creates a spreadsheet with time spent per project using Google mail and calendar data

## Setup Instructions

### 1. Create a new Google Cloud Platform (GCP) project
------------------------------------------------

To use Google Workspace APIs, you need a Google Cloud Platform project. This project forms the basis for creating, enabling, and using all GCP services, including managing APIs, enabling billing, adding and removing collaborators, and managing permissions.

1.  Open the [Google Cloud Console](https://console.cloud.google.com/).
2.  Next to "Google Cloud Platform," click the Down arrow arrow_drop_down . A dialog listing current projects appears.
3.  Click New Project. The New Project screen appears.
4.  In the Project Name field, enter a descriptive name for your project.
5.  Click Organization and select your organization.
6.  In the Location field, click Browse to display potential locations for your project.
7.  Click a location and click Select.
8.  Click Create. The console navigates to the Dashboard page and your project is created within a few minutes.

### 2. Enable a Google Workspace API
-----------------------------

1.  Open the [Google Cloud Console](https://console.cloud.google.com/).
2.  Next to "Google Cloud Platform," click the Down arrow arrow_drop_down and select a project.
3.  In the top-left corner, click Menu menu > APIs & Services.
4.  Click Enable APIs and Services. The Welcome to API Library page appears.
5.  In the search field, enter the name of the API you want to enable. **In this case**, "Calendar API" to find the Google Calendar API.
6.  Click the API to enable. The API page appears.
7.  Click Enable. The Overview page appears.

### 3. Obtain required crendentials for API Usage
-----------------------------

1. Open the [Google Cloud Console](https://console.cloud.google.com/).
2. Next to "Google Cloud Platform," click the Down arrow arrow_drop_down and select a project.
3. In the top-left corner, click Menu menu > APIs & Services > Credentials.
4. Click on "+ Create credentials" on the top menu and select `OAuth client ID`
5. On Application type, select Web Application
6. Choose a name of your preference and click on Create
8. The credentials page will appear again. Search for your newly created credentials in the OAuth 2.0 Client IDs
9. Click on the Download icon at the right of your credentials name
10. Change it's name to `credentials` and store it on the same folder of this script
