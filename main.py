import imaplib
import smtplib
from email.message import EmailMessage
import imghdr
import threading
from time import time, sleep
import json
import jwt
import requests
from bs4 import BeautifulSoup
import datetime
import pickle
import os.path
import http.client as http
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# To install all required modules: pip install -r requirements.txt
# If on linux use pip3


class Gmail:

    def __init__(self, username: str, password: str, imap: str = 'imap.gmail.com',
                 port: int = 993, timeout: int = 5, look_pretty: bool = True):
        self.username = username
        self.password = password
        self.imap = imap
        self.port = port
        self.timeout = timeout
        self.look_pretty = look_pretty
        self.connection = imaplib.IMAP4_SSL(self.imap, self.port)
        self.login()

    def login(self):
        """
        This function logs in with the credentials provided. This is outside of the init function
        just in case you are logged out you can call on this function.
        :return: None
        """
        self.connection.login(self.username, self.password)

    def get_body(self, data) -> list:
        """
        This function looks through the data that was provided by the get_messages function and
        return a list of emails in text format. This should not be run alone.
        :param data:
        :return:
        """
        result = []
        for messages in data[::-1]:
            for message in messages:
                if type(message) is tuple:
                    if self.look_pretty:
                        result.append(BeautifulSoup(message[1], 'lxml').getText()
                                      .replace('=E2=80=A2', '-').replace('=E2=80=99', '\''))
                    else:
                        result.append(message[1])
        return result

    def set_label(self, label: str):
        """
        This function tells the email gatherer where to send its search party.
        :param label:
        :return:
        """
        return self.connection.select(label)

    def get_messages(self, search: str = '(ALL)') -> list:
        """
        This function searches and checks for any emails within the search parameters.
        :param search:
        :return:
        """
        _, result = self.connection.search(None, search)
        data = []
        for num in result[0].split():
            _, msg = self.connection.fetch(num, '(RFC822)')
            data.append(msg)
        return self.get_body(data)


class Alert:

    def __init__(self, username, password):
        self.username = username
        self.password = password
        self.MMS = [
            '@mms.att.net',  # at&t/Cricket
            '@tmomail.net',  # T-Mobile
            '@vzwpix.com',  # Verizon Wireless
            # '@pm.sprint.com', # Sprint
            '@mypixmessages.com',  # XFinity
            '@vmpix.com',  # Virgin Mobile
            # '@msg.fi.google.com',  # Google Fi
            '@mmst5.tracfone.com'  # Tracfone
        ]

    def SendAlert(self, Receiver: str, Message: str, Subject: str = None, Image: str = None,
                  From: str = None):
        """
        This function is designed to receive a emails address or phone number and send it a message.
        :param Receiver:
        :param Message:
        :param Subject:
        :param Image:
        :param From:
        :return:
        """
        if Receiver.find('@') != -1:
            self.Email(Receiver, Message, Subject, Image, From)
        else:
            self.Text(Receiver, Message, Subject, Image, From)

    def Text(self, Receiver: str, Message: str, Subject: str = None, Image: str = None, From: str = None):
        """
        This function takes a phone number and sends it a message from an email.
        :param Receiver:
        :param Message:
        :param Subject:
        :param Image:
        :param From:
        :return:
        """
        Messages = []
        for Provider in self.MMS:
            msg = EmailMessage()
            msg['Subject'] = Subject
            msg['From'] = From
            msg['To'] = Receiver + Provider
            msg.set_content(Message)

            if Image:
                with open(Image, 'rb') as f:
                    file_data = f.read()
                    file_type = imghdr.what(f.name)
                    file_name = f.name
                msg.add_attachment(file_data, maintype='image', subtype=file_type, filename=file_name)
            Messages.append(msg)

        for m in Messages:
            threading.Thread(target=self.Send, args=(m,)).start()

    def Email(self, Receiver: str, Message: str, Subject: str = None, Image: str = None, From: str = None,
              Cc: str = None):
        """
        This function sends an email. Passing a From argument can mask your email address.
        :param Cc:
        :param Receiver:
        :param Message:
        :param Subject:
        :param Image:
        :param From:
        :return:
        """
        msg = EmailMessage()
        msg['Subject'] = Subject
        msg['From'] = From
        msg['To'] = Receiver
        msg['Cc'] = Cc
        msg.set_content(Message)

        if Image:
            with open(Image, 'rb') as f:
                file_data = f.read()
                file_type = imghdr.what(f.name)
                file_name = f.name
            msg.add_attachment(file_data, maintype='image', subtype=file_type, filename=file_name)
        threading.Thread(target=self.Send, args=(msg,)).start()

    def Send(self, msg):
        """
        This function sends any messages that have been compiled by the functions above. This should
        not be run alone.
        :param msg:
        :return:
        """
        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(self.username, self.password)
                smtp.send_message(msg)
            return True
        except Exception as e:
            print(e)
            return False


class Zoom:

    def __init__(self, key: str, secret: str):
        self.key = key
        self.secret = secret

    def create_auth_token(self, algorithm: str = 'HS256') -> jwt.encode:
        """
        This function generates a javascript token with the zoom key and secret. This should
        not be run alone.
        :param algorithm:
        :return:
        """
        return jwt.encode({'iss': self.key, 'exp': time() + 5000}, self.secret, algorithm=algorithm)

    def create_meeting(self, data, duration: str = '60', timezone: str = 'America/Los_Angeles') -> dict:
        """
        This function takes a event object and creates a zoom link with the data within. If you wish
        to add to this function the link to the api is here:
        https://marketplace.zoom.us/docs/api-reference/zoom-api/meetings/meeting
        :param data:
        :param duration:
        :param timezone:
        :return:
        """
        token = self.create_auth_token()
        meeting_details = {'topic': f'Tutoring Session With {data.invitee}',
                           'type': 2,
                           'start_time': f"{data.date}T{data.time}Z",
                           'duration': duration,
                           'timezone': timezone,
                           'settings': {'host_video': 'true',
                                        'participant_video': 'true',
                                        'waiting_room': 'true',
                                        }
                           }
        headers = {'authorization': f'Bearer {token}',
                   'content-type': 'application/json'}
        r = requests.post(f'https://api.zoom.us/v2/users/me/meetings', headers=headers,
                          data=json.dumps(meeting_details))
        return json.loads(r.text)


class Event:

    def __init__(self, message: str):
        scraped = self.scrap_info(message)
        if not scraped:
            self.valid = False
        else:
            self.valid = True
            self.invitee = scraped['Invitee']
            self.invitee_email = scraped['Invitee Email']
            self.invitee_timezone = scraped['Invitee Time Zone']
            self.time = scraped['Time']
            self.date = scraped['Date']
            self.notified = False
            self.zoom_id = None
            self.zoom_join_url = None
            self.zoom_passcode = None
        del scraped

    def __str__(self):
        """
        This function is called upon if you try to print this class object. It will format the
        response so its easier to read.
        :return:
        """
        return f'\n' \
               f'Invitee: {self.invitee}\n' \
               f'Invitee Email: {self.invitee_email}\n' \
               f'Invitee Timezone: {self.invitee_timezone}\n' \
               f'Event Start: {self.time} {self.date}\n' \
               f'Zoom ID: {self.zoom_id}\n' \
               f'Zoom Meeting Url: {self.zoom_join_url}\n' \
               f'Zoom Passcode: {self.zoom_passcode}\n' \
               f'Notified: {str(self.notified)}'

    def scrap_info(self, text: str, start: str = 'Event Type:', end: str = 'View event in Calendly'):
        """
        This function takes the raw data of the emails and converts it into data we can use. This
        should not be run alone.
        :param text:
        :param start:
        :param end:
        :return:
        """
        if start not in text and end not in text:
            return None
        months = {'january': '01', 'february': '02', 'march': '03', 'april': '04', 'may': '05', 'june': '06',
                  'july': '07',
                  'august': '08', 'september': '09', 'october': '10', 'november': '11', 'december': '12'}
        result = {}
        updated = '\n'.join([m.rstrip('=') for m in text.split('\r\n')])
        updated = updated[updated.find(start):updated.find(end) + len(end)].split('\n')
        for i, entry in enumerate(updated):
            if 'Invitee:' in entry:
                result['Invitee'] = updated[i + 1]
            if 'Invitee Email:' in entry:
                result['Invitee Email'] = updated[i + 1]
            if 'Event Date/Time' in entry:
                mtime = updated[i + 1].split('-')[0].strip()
                result['Time'] = str(int(mtime[:2]) + 12) + mtime[2:5] + ':00' if 'pm' in mtime and mtime[:2] != '12' \
                    else mtime[:5] + ':00'
                date = updated[i + 1].split(',')[1].split('(')[0].strip().split()
                date[0] = '0' + date[0] if len(date[0]) == 1 else date[0]
                result['Date'] = date[2] + '-' + months[date[1].lower()] + '-' + date[0]
            if 'Invitee Time Zone:' in entry:
                result['Invitee Time Zone'] = updated[i + 1].split('-')[0].strip()
        return result

    def convert_date(self) -> datetime.date:
        """
        This function converts the string date to a datetime.date object so it can
        be compared to other datetime.date objects.
        :return:
        """
        event_date = self.date.split('-')
        return datetime.datetime(int(event_date[0]), int(event_date[1]), int(event_date[2])).date()

    def get_day_of_week(self) -> str:
        """
        This function converts the date of the event to the day of the week.
        :return:
        """
        return self.convert_date().strftime('%A')

    def get_month(self) -> str:
        """
        This function converts the date of the event to a month of the year.
        :return:
        """
        return self.convert_date().strftime('%B')

    def get_suffix(self) -> str:
        """
        This function return the day with a suffix appended to it.
        :return:
        """
        day = self.convert_date().day
        suffix = 'th' if 11 <= day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
        return str(day) + suffix

    def get_standard_time(self) -> str:
        """
        This function converts military time back to standard time.
        :return:
        """
        return datetime.datetime.strptime(self.time, '%H:%M:%S').strftime('%I:%M %p')

    def end_time(self):
        return (datetime.datetime.strptime(self.time, '%H:%M:%S') + datetime.timedelta(minutes=50)).time()


class Calendar:

    def __init__(self, email: str, user_file: str = 'calendar_auth.json'):
        self.user_file = user_file
        self.user = self.login()
        self.user_email = email

    def login(self):
        """
        This function will check for credentials and create new ones if there are none
        on file.
        :return:
        """
        cred = None
        scopes = ['https://www.googleapis.com/auth/calendar']
        if os.path.isfile(self.user_file):
            cred = Credentials.from_authorized_user_file(self.user_file, scopes)
        if not cred or not cred.valid:
            if cred and cred.expired and cred.refresh_token:
                cred.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', scopes)
                cred = flow.run_local_server(port=0)
            with open(self.user_file, 'w') as f:
                f.write(cred.to_json())

        return build('calendar', 'v3', credentials=cred)

    def add_calendar_event(self, data, timezone: str = 'America/Los_Angeles'):
        """
        This function creates the google calendar events.
        :param data:
        :param timezone:
        :return:
        """
        event = {
            'summary': f'Tutoring Event With {data.invitee}!',
            'description': f'Meeting with {data.invitee}.\n'
                           f'Email Address: {data.invitee_email}\n'
                           f'Invitee timezone: {data.invitee_timezone}\n'
                           f'Zoom Link: {data.zoom_join_url}',
            'start': {
                'dateTime': f'{data.date}T{data.time}',
                'timeZone': timezone,
            },
            'end': {
                'dateTime': f'{data.date}T{data.end_time()}',
                'timeZone': timezone,
            },
            'attendees': [
                {'email': data.invitee_email},
                {'email': self.user_email}
            ]
        }
        self.user.events().insert(calendarId='primary', body=event).execute()


class Database:

    def __init__(self):
        self.event_data = []
        self.last_clean_up = datetime.date.today()

    def append_event(self, event, zoom: Zoom, calendar, outdated_ok: bool = False) -> bool:
        """
        This function is the gateway to add events into the database. It will confirm their
        validity and make sure its not adding any old events. This is also where a zoom link
        is created and where the event is added to the google calendar.
        :param event:
        :param zoom:
        :param calendar:
        :param outdated_ok:
        :return:
        """
        if not event.valid:
            return False
        if event.convert_date() < datetime.date.today() and not outdated_ok:
            return False
        if self.event_saved(event):
            return False
        if not event.zoom_join_url:
            meeting = zoom.create_meeting(event)
            event.zoom_id = meeting['id']
            event.zoom_join_url = meeting['join_url']
            event.zoom_passcode = meeting['password']
        # This is the line that adds google calendar events.
        calendar.add_calendar_event(event)
        print(f'{event.invitee} has been scheduled for {event.get_month()} {event.get_suffix()}.')
        self.event_data.append(event)
        return True

    def event_saved(self, questioned) -> bool:
        """
        This function checks if the event in question is already in the database so
         it does not add any duplicates to the system.
        :param questioned:
        :return:
        """
        for event in self.event_data:
            if (event.invitee, event.date, event.time) == \
                    (questioned.invitee, questioned.date, questioned.time):
                return True
        return False

    def events_on_date(self, date: datetime.date) -> list:
        """
        This function returns any events that are happening on the specified date.
        :param date:
        :return:
        """
        data = []
        for event in self.event_data:
            if event.convert_date() == date:
                data.append(event)
        return data

    def events_before_date(self, date: datetime.date, indexed: bool = False) -> list:
        """
        This function return any events before a certain date with an option to return the
        index values of those events.
        :param date:
        :param indexed:
        :return:
        """
        data = []
        for count, event in enumerate(self.event_data):
            if event.convert_date() < date:
                if indexed:
                    data.append(count)
                else:
                    data.append(event)
        return data

    def events_after_date(self, date: datetime.date, indexed: bool = False) -> list:
        """
        This function return any events after a certain date with an option to return the
        index values of those events.
        :param date:
        :param indexed:
        :return:
        """
        data = []
        for count, event in enumerate(self.event_data):
            if event.convert_date > date:
                if indexed:
                    data.append(count)
                else:
                    data.append(event)
        return data

    def events_tomorrow(self) -> list:
        """
        This function returns a list of events that are upcoming tomorrow.
        :return:
        """
        return self.events_on_date(datetime.date.today() + datetime.timedelta(days=1))

    def events_today(self) -> list:
        """
        This function returns a list of events that are happening today.
        :return:
        """
        return self.events_on_date(datetime.date.today())

    def events_yesterday(self) -> list:
        """
        This function return any events that happened yesterday.
        :return:
        """
        return self.events_on_date(datetime.date.today() - datetime.timedelta(days=1))

    def events_upcoming(self) -> list:
        """
        This function returns any upcoming events.
        :return:
        """
        return self.events_after_date(datetime.date.today() - datetime.timedelta(days=1))

    def list_unprepared(self, days_ahead: int = 1):
        """
        This function will gather anyone that has not been sent a preparation email.
        :param days_ahead: How many days ahead would you like too notify.
        :return:
        """
        result = []
        for i in range(days_ahead + 1):
            potential = self.events_on_date(datetime.date.today() + datetime.timedelta(days=i))
            for event in potential:
                if not event.notified:
                    result.append(event)
        return result

    def remove_event(self, event: Event) -> bool:
        try:
            del self.event_data[self.event_data.index(event)]
            return True
        except ValueError:
            return False

    def mark_as_notified(self, event: Event) -> bool:
        """
        This function will search for the event provided and mark it as a successfully
        emailed participant.
        :param event:
        :return:
        """
        try:
            self.event_data[self.event_data.index(event)].notified = True
            return True
        except ValueError:
            return False

    def search_by_name(self, name: str) -> list:
        """
        This function takes a name and looks through all the events to return a
        list of events that have that name on file.
        :param name:
        :return:
        """
        result = []
        for event in self.event_data:
            if name.lower() in event.invitee.lower():
                result.append(event)
        return result

    def search_by_name_date_time(self, name: str, date: str, time_: str):
        for event in self.event_data:
            if (event.invitee, event.date, event.time) == (name, date, time_):
                return event
        return None

    def check_for_cleanup(self) -> None:
        """
        This function checks if the database has been cleaned today and if not
        then it will call on the self cleanup function.
        :return:
        """
        if self.last_clean_up < datetime.date.today():
            self.cleanup()

    def cleanup(self) -> int:
        """
        This function checks for any events that have passed and will delete them.
        :return:
        """
        old_events = self.events_before_date(datetime.date.today(), indexed=True)
        for count, i in enumerate(old_events):
            del self.event_data[i - count]
        self.last_clean_up = datetime.date.today()
        return len(old_events)

    def destroy(self) -> None:
        """
        This function removes all the event data saved within the database.
        :return:
        """
        if 'y' == input("Would you like to delete the contents of the database? (y/n): ").lower():
            del self.event_data[:]
            save_db(self)
            print('Data deleted.')


def save_db(db: Database, file: str = 'tutoring.database') -> None:
    """
    This function takes a Database object and serializes the data so it can be
    saved as a file and loaded again at a later date.
    :param db:
    :param file:
    :return:
    """
    with open(file, 'wb') as f:
        pickle.dump(db, f)


def open_db(file: str = 'tutoring.database') -> Database:
    """
    This function returns a database object, if there is not one on file then it
    will create a new one.
    :param file:
    :return:
    """
    if os.path.isfile(file):
        with open(file, 'rb') as f:
            saved_db = pickle.load(f)
        return saved_db
    return Database()


def internet_active() -> bool:
    """
    This function checks if you are connected to the internet. It only request the head of
    the website as to not waist time gathering the whole page.
    """
    connection = http.HTTPConnection("www.google.com", timeout=5)
    try:
        connection.request("HEAD", "/")
        connection.close()
        return True
    except:
        connection.close()
        return False


def main():
    # Setting the database as global so it can be called upon within the python interactive mode.
    global database

    # Waits for an internet connection.
    while not internet_active():
        sleep(5)

    # Opens the database / create new one if none exists.
    database = open_db()
    # Cleans up database to save disk space. Commenting this out will not affect the program.
    print(database.check_for_cleanup(), 'Old Events Cleared.')

    # Logging into Gmail with 'username' and 'password' credentials. You will either have to
    # set up a app key or make your account unsecure to allow the program to read your emails.
    # Here is the link: https://myaccount.google.com/security
    gmail_retriever = Gmail('YOUR_GMAIL_USERNAME', 'YOUR_GMAIL_PASSWORD')
    gmail_sender = Alert('YOUR_GMAIL_USERNAME', 'YOUR_GMAIL_PASSWORD')

    # Logging into Zoom with 'key' and 'secret' credentials. Here is a tutorial on the zoom
    # api: https://www.geeksforgeeks.org/how-to-create-a-meeting-with-zoom-api-in-python/
    zoom = Zoom('YOUR_ZOOM_KEY', 'YOUR_ZOOM_SECRET')

    # Logging into google calendar, the first time it will prompt you to give it your
    # credentials but will then save a token that can automate the process in the future.
    # if you do not have a credentials.json file you will have to activate the
    # "Google Calendar API" and register as a google developer account to enable this
    # program to edit and add events to your calendar by downloading the credentials you've
    # created and store them in the same directory as this program.
    # Here is the link: https://console.cloud.google.com/
    calendar = Calendar('YOURGMAIL@gmail.com')

    # Telling Gmail to look at anything set with the "Tutoring" Label.
    gmail_retriever.set_label('Tutoring')

    # Gather all emails and loop through them.
    for message in gmail_retriever.get_messages():
        # Gather only the useful information from the message and send it to the database
        event = Event(message)
        database.append_event(event, zoom, calendar)

    for event in database.list_unprepared():
        subject = f'Cybersecurity Boot Camp - Tutorial Confirmation - {event.get_day_of_week()}, {event.get_month()} ' \
                  f'{event.get_suffix()}, at {event.get_standard_time()}, Pacific.'

        message = f'Hi {event.invitee.split(" ")[0]}!\n' \
                  f'Thank you for scheduling your session with me. I am looking forward to our session on ' \
                  f'{event.get_day_of_week()}, {event.get_month()} {event.get_suffix()}, at ' \
                  f'{event.get_standard_time()}, Pacific.\n\n' \
                  f'If something comes up and the scheduled time will not work, let me know a minimum of 6 hours' \
                  f' before the appointment time and we’ll figure something out.\n\n' \
                  f'This session will take place here:\n\n' \
                  f'Join Zoom Meeting\n' \
                  f'{event.zoom_join_url}\n\n' \
                  f'Meeting ID: {event.zoom_id}\n' \
                  f'Passcode: {event.zoom_passcode}\n\n' \
                  f'(If you have not used zoom before please join the meeting at least 15 minutes early because it' \
                  f' may have you download and install some software.)\n\n' \
                  f'Again, all I need from you:\n' \
                  f'• Be on Tutors & Students Slack 5 minutes before your time slot.\n' \
                  f'• Make sure your computer/mic/internet connection is working.\n' \
                  f'• Make sure your workspace is quiet and free from interruptions.\n' \
                  f'• At the end of the session, I will provide you with a link to a 2-minute evaluation form ' \
                  f'that you are required to complete.\n\n' \
                  f'Slack or email me with any questions. I’m looking forward to our meeting!\n\n' \
                  f'Please Reply All to this email so that I know you have seen it.\n\n' \
                  f'(CC Central Support on all tutor emails by always using REPLY ALL).\n\n' \
                  f'Sincerely,\n' \
                  f'Nathan Barnes\n'
        gmail_sender.Email(Receiver=event.invitee_email, Message=message, Subject=subject, Image=None,
                           Cc='centraltutorsupport@bootcampspot.com')
        print(f'{event.invitee} is receiving an email!')
        database.mark_as_notified(event)

    # serialize the database and save it as a file.
    save_db(database)


if __name__ == '__main__':
    main()
