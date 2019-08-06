#!/usr/bin/env python3
from __future__ import print_function
# Built In Modules
import datetime
import os
import os.path

# 3rd Party Modules
import requests
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# Global Variables
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = os.environ['SPREADSHEET_ID']


def get_google_sheets():
    """
    Gets the Sheets client
    """
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
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    return sheet


def update_google_sheets(client, values):
    """
    Function to update the Google Sheet.


    :param googleapiclient.discovery.Resource client: Sheets client
    :param Dict values: Values to add
    """
    # Loop through the date column until we find a the right place to add a column
    today = datetime.datetime.today()
    index = 2
    while True:
        # Find the appropriate row to append
        date_compare = get_values(client, 1, 1, index, index)
        date_compare = datetime.datetime.strptime(date_compare[0][0], '%Y-%m-%d')
        if today < date_compare:
            index -= 1
            break
        index += 1
    # Once we have the row we can add the needed values
    # First add a new row
    client.batchUpdate(spreadsheetId=SPREADSHEET_ID,
                       body={
                           'requests': [
                               {'insertDimension': {
                                   'inheritFromBefore': True,
                                   'range': {'endIndex': index+1,
                                             'startIndex': index,
                                             'dimension': 'ROWS',
                                             'sheetId': 0}
                                     }
                                }
                           ]}).execute()
    # Now add our values to the new row
    # First format out data
    remaining_work = f"=SUM(B{index+1}:C{index+1})"
    total_work = f"=SUM(B{index+1}:D{index+1})"
    daily_velocity = f"=TRUNC((D{index+1}-D{index})/(A{index+1}-{index}), 1)"
    daily_scope = f"=TRUNC((F{index+1}-F{index}) / (A{index+1}-A{index}), 1)"
    vals = [[values['date'], values['todo'], values['in_progress'], values['done'],
             remaining_work, total_work, daily_velocity, daily_scope]]

    # Now add the values
    client.values().append(spreadsheetId=SPREADSHEET_ID,
                           range=get_coordinates_string(1, 8, index + 1, index + 1),
                           body={'values': vals},
                           valueInputOption='USER_ENTERED').execute()
    # Now format our values
    client.batchUpdate(spreadsheetId=SPREADSHEET_ID,
                       body={
                           'requests': [
                               {
                                   'repeatCell': {
                                       'range': {
                                           'sheetId': 0,
                                           'startRowIndex': index+1,
                                           'endRowIndex': index+1
                                       },
                                       'cell': {
                                           'horizontalAlignment': 'RIGHT'
                                       }
                                   }
                               }
                           ]
                       })


def get_values(client, x1, x2, y1, y2, sheet='Sheet1'):
    """
    Helper function to get values from sheet.


    :param googleapiclient.discovery.Resource client: Sheets client
    :param Int x1: X1 coordinate
    :param Int x2: X2 coordinate
    :param Int y1: Y1 coordinate
    :param Int y2: Y2 coordinate
    :param String sheet: Which sheet to parse data from (default is Sheet1)
    :return: Values
    :rtype: List
    """
    pass
    if sheet == 'Sheet1':
        resp = client.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=get_coordinates_string(x1, x2, y1, y2)).execute()
    else:
        resp = client.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=get_coordinates_string(x1, x2, y1, y2, sheet=sheet)).execute()
    return resp.get('values', [])


def get_coordinates_string(x1, x2, y1, y2, sheet='Sheet1'):
    """
    Helper function to return the A1 (Google Sheet) notation.
    :param Int x1: X1 coordinate
    :param Int x2: X2 coordinate
    :param Int y1: Y1 coordinate
    :param Int y2: Y2 coordinate
    :param String sheet: Which sheet to parse data from (default is Sheet1)
    :return: A1 notation
    :rtype: String
    """
    return f"{sheet}!{get_letter_from_coordinate(x1)}{y1}:{get_letter_from_coordinate(x2)}{y2}"


def get_letter_from_coordinate(x):
    """
    Helper function to get coordinate from number


    :param Int x: X coordinate
    :return: Letter corresponding to number
    :rtype: String
    """
    alpha = ['INVALID', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K',
             'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
             'V', 'W', 'X', 'Y', 'Z']
    return alpha[x]


def main():
    """
    Main function to start the program.
    """
    # First get our JIRA data
    values = get_data()

    # Get the Sheets client
    client = get_google_sheets()

    # Update the appropriate value
    update_google_sheets(client, values)



def get_data():
    """
    Helper function to get data from JIRA
    :return: Values
    :rtype: Dict
    """
    auth = (os.environ['JIRA_USER'], os.environ['JIRA_PW'])
    url = 'https://projects.engineering.redhat.com/rest/api/2/search'
    kwargs = {
        'server': 'https://projects.engineering.redhat.com',
        'options': {'basic_auth': auth}
    }
    values = [datetime.date.today().strftime("%Y-%m-%d")]
    for idx in (30032, 30033, 30034):
        params = dict(
            jql="filter=%i" % idx,
            validateQuery=True,
            startAt=0,
            maxResults=10,
        )
        response = requests.get(url, params=params, auth=auth)
        values.append(str(response.json()['total']))
    return {
        'date': values[0],
        'todo': values[1],
        'in_progress': values[2],
        'done': values[3]
    }
    import pdb; pdb.set_trace()


if __name__ == '__main__':
    main()
