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
	# This code add 4kurs fiit subjects to my Google Calendar 
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('client_id.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=creds.authorize(Http()))

    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time

    # Connect to XML file and search 4 kurs FIIT-15
    book = xlrd.open_workbook('C:/1/l01gcal-poewmuh/imi2018.xls')
    sh = book.sheets()[8]

    start = ['08:00', '9:50', '11:40', '14:00', '15:40', '17:30']
    end =  ['09:35', '11:25', '13:15', '15:35', '17:20', '19:05']

    #day
    j = 7

    for i in range(3,39):
    	npara = (i-3)%6
    	if npara==0:
    		j += 1
    	if sh.cell(i,8).value != '':
            event = {
                'summary': sh.cell(i,8).value,
                'location': sh.cell(i,10).value,
                'description': sh.cell(i,9).value,
                'start': {
                    'dateTime': '2018-10-{:02d}T{}:00+09:00'.format(j, start[npara]),
                    'timeZone': 'Asia/Yakutsk',
                },
                'end': {
                    'dateTime': '2018-10-{:02d}T{}:00+09:00'.format(j, end[npara]),
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