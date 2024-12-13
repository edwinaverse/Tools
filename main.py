#!/bin/bash

"""
Owner: Edwina E.

Purpose: Update School calendars on Google Calendar because I'm lazy.
"""

import datetime
import os.path
import pickle
import pandas as pd
import csv
import copy

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/calendar']
TIMED_EVENT = {
        'summary': None,
        'location': None,
        'description': None,
        'start': {
            'dateTime': None,
            'timeZone': 'America/New_York',
        },
        'end': {
            'dateTime': None,
            'timeZone': 'America/New_York',
        }
}

ALL_DAY_EVENT = {
'summary': None,
        'location': None,
        'description': None,
        'start': {
            'date': None,
            'timeZone': 'America/New_York',
        },
        'end': {
            'date': None,
            'timeZone': 'America/New_York',
        }
}

def main():
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
                '/Users/edwina/Downloads/certificate.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    # In this case, I'll be creating a new calendar
    new_calendar = {
        'summary': 'School',
        'timeZone': 'America/New_York'
    }
    created_calendar = service.calendars().insert(body=new_calendar).execute()
    print(f"Created calendar: {created_calendar['id']}")

    # excel -> csv
    tempcal = pd.read_excel("/Users/edwina/Downloads/schoolcalendar.xlsx")
    # Write the dataframe to a CSV file
    calendar = tempcal.to_csv("calendar.csv", index=False)

    tempEvent = None
    with open("calendar.csv") as calendar:
        calendarIter = csv.reader(calendar)
        # skips the headers
        next(calendarIter, None)
        for row in calendarIter:
            tempEvent = copy.deepcopy(ALL_DAY_EVENT)
            tempEvent['summary'] = str(row[3])
            tempEvent['location'] = None
            tempEvent['description'] = str(row[10])
            tempEvent['start']['date'] = fix_date(row[6])
            tempEvent['end']['date'] = fix_date(row[8])
            created_event = service.events().insert(calendarId=created_calendar['id'], body=tempEvent).execute()
            print(f"Created event: {created_event['id']}")

def fix_date(origDate):
    """ Translates Month/Day/Year format to iso format for Google Calendar """
    month, day, year = origDate.split('/')
    origDate = datetime.datetime(int(year), int(month), int(day))
    return origDate.isoformat().split("T")[0]


if __name__ == '__main__':
    main()