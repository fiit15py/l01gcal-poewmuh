from __future__ import print_function
import datetime
import os
import xlrd
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/calendar'

def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('client_id.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=creds.authorize(Http()))


    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    book = xlrd.open_workbook('C:/1/l01gcal-poewmuh/imi2018.xls')
    sh = book.sheets()[8]
    start = ['08:00', '9:50', '11:40', '14:00', '15:40', '17:30']
    end =  ['09:35', '11:25', '13:15', '15:35', '17:20', '19:05']
    for i in range(38):
    	if sh.cell(i,8).value != '':
            event = {
                'summary': sh.cell(i,8).value,
                'location': sh.cell(i,10).value,
                'description': sh.cell(i,9).value,
                'start': {
                    'dateTime': '2018-10-08T'+ start[(i-3) % 6] +':00+09:00',
                    'timeZone': 'Asia/Yakutsk',
                },
                'end': {
                    'dateTime': '2018-10-08T'+ end[(i-3) % 6] +':00+09:00',
                    'timeZone': 'Asia/Yakutsk',
                },
                'recurrence': [
                    'RRULE:FREQ=WEEKLY;COUNT=12'
                ],
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                     {'method': 'popup', 'minutes': 30},
                     ],
                }
                }

            event = service.events().insert(calendarId='primary', body=event).execute()
            print('Event created: %s' % (event.get('htmlLink')))

if __name__ == '__main__':
    main()