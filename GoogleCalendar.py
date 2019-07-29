from datetime import datetime, timedelta
import pickle
import os.path
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ["https://www.googleapis.com/auth/calendar"]

CREDENTIALS_FILE = "C:/Users/gmcwilliams/OneDrive/Documents/personal/apps/client_secret_py-google-calendar.json"

class GoogleCalendar:

    tz = "Europe/London"
    service = None

    def __init__(self, calendarName: str):
        creds = None
        if self.service is None:
            creds = None
            # The file token.pickle stores the user's access and refresh tokens, and is
            # created automatically when the authorization flow completes for the first
            # time.
            if os.path.exists("token.pickle"):
                with open("token.pickle", "rb") as token:
                    creds = pickle.load(token)
            # If there are no (valid) credentials available, let the user log in.
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        CREDENTIALS_FILE, SCOPES
                    )
                    creds = flow.run_local_server(port=0)

                # Save the credentials for the next run
                with open("token.pickle", "wb") as token:
                    pickle.dump(creds, token)

            self.service = build("calendar", "v3", credentials=creds)
            self.get_named_calendar_(calendarName)


    def get_named_calendar_(self, calendarId: str):
        print("Getting list of calendars")
        calendars_result = self.service.calendarList().list().execute()

        calendars = calendars_result.get("items", [])

        if not calendars:
            print("No calendars found.")
        for calendar in calendars:
            if calendar["summary"] == calendarId:
                summary = calendar["summary"]
                self.namedCalendarId = calendar["id"]
                self.tz = calendar["timeZone"]
                primary = "Primary" if calendar.get("primary") else ""
                print("%s\t\t%s\t%s" % (summary, self.namedCalendarId, primary))


    def create_event(self):
        today = datetime.now().date()
        tomorrow = datetime(today.year, today.month, today.day, 10) + timedelta(days=1)
        start = tomorrow.isoformat()
        end = (tomorrow + timedelta(hours=1)).isoformat()

#                    "id": "icalXyzzyXyz",
#                    "iCalUID": "automating",
        self.event = (
            self.service.events()
            .insert(
                calendarId=self.namedCalendarId,
                body={
                    "summary": "Automating calendar",
                    "description": "This is a tutorial example of automating google calendar with python",
                    "start": {"dateTime": start, "timeZone": self.tz},
                    "end": {"dateTime": end, "timeZone": self.tz},
                },
            )
            .execute()
        )

        print("created event")
        print(f'id: {self.event["id"]}/{self.event["iCalUID"]}]')
        print("summary: ", self.event["summary"])
        print("starts at: ", self.event["start"]["dateTime"])
        print("ends at: ", self.event["end"]["dateTime"])


    def delete_eventId_(self, eventId: str):
        try:
            self.service.events().delete(calendarId=self.namedCalendarId, eventId=eventId).execute()
            print(f"Event deleted:{eventId}")
        except HttpError:
            print(f"Failed to delete event:{eventId}")


    def delete_eventId(self, eventId: str):
        self.service.events().delete(calendarId=self.namedCalendarId, eventId=eventId).execute()
        print(f"Event deleted:{eventId}")


    def delete_event(self):
        delete_eventId(self.event["id"])


    def list_calendars(self):
        # Call the Calendar API
        print("Getting list of calendars")
        calendars_result = self.service.calendarList().list().execute()

        calendars = calendars_result.get("items", [])

        if not calendars:
            print("No calendars found.")
        for calendar in calendars:
            summary = calendar["summary"]
            id = calendar["id"]
            primary = "Primary" if calendar.get("primary") else ""
            print("%s\t\t%s\t%s" % (summary, id, primary))


    def get_events(self):
        self.get_events_from(datetime.utcnow())


    def get_events_from(self, start: datetime):
        fromdate = start.isoformat() + "Z"  # 'Z' indicates UTC time
        print("Getting events from:", fromdate)
        events_result = (
            self.service.events()
            .list(
                calendarId=self.namedCalendarId,
                timeMin=fromdate,
                singleEvents=True,
                orderBy="startTime",
                showDeleted="true",
            )
            .execute()
        )
        self.events = events_result.get("items", [])

        if not self.events:
            print("No upcoming events found.")
        for event in self.events:
            start = event["start"].get("dateTime", event["start"].get("date"))
            print(start, event["summary"], event["id"], event["iCalUID"], event["status"])


def main():
    google_calendar = GoogleCalendar("py-google-calendar")
#    google_calendar.list_calendars()
    google_calendar.get_events()
    google_calendar.create_event()
#    google_calendar.delete_event()
#    google_calendar.delete_eventId("automating")

if __name__ == "__main__":
    main()
