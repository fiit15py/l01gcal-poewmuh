from __future__ import print_function
import datetime
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

    event = {
      'summary': 'ДВ. Администрирование OC Windows Леверьев В.С.',
      'location': '430',
      'description': 'л / пр',
      'start': {
        'dateTime': '2018-10-08T15:45:00+09:00',
        'timeZone': 'Asia/Yakutsk',
      },
      'end': {
        'dateTime': '2018-10-08T17:20:00+09:00',
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