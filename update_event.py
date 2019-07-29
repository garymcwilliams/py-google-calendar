from datetime import datetime, timedelta
from cal_setup import get_calendar_service


def main():
    # update the event to tomorrow 9 AM IST
    service = get_calendar_service()

    d = datetime.now().date()
    tomorrow = datetime(d.year, d.month, d.day, 9)+timedelta(days=1)
    start = tomorrow.isoformat()
    end = (tomorrow + timedelta(hours=2)).isoformat()

    event_result = service.events().update(
      calendarId='primary',
      eventId='njdev79d574rdmkv0180t7t7lo',
      body={
       "summary": 'Updated Automating calendar',
       "description": 'This is a tutorial example of automating google calendar with python, updated time.',
       "start": {"dateTime": start, "timeZone": 'Europe/London'},
       "end": {"dateTime": end, "timeZone": 'Europe/London'},
       },
    ).execute()

    print("updated event")
    print("id: ", event_result['id'])
    print("summary: ", event_result['summary'])
    print("starts at: ", event_result['start']['dateTime'])
    print("ends at: ", event_result['end']['dateTime'])

if __name__ == '__main__':
    main()