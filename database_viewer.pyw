from main import *
import main
import PySimpleGUI as sg


class Event:

    def __init__(self, name, email, time_, date, notified, timezone, zoom_url, zoom_passcode, zoom_id):
        self.valid = True
        self.invitee = name
        self.invitee_email = email
        self.invitee_timezone = timezone
        self.time = time_
        self.date = date
        self.notified = notified
        self.zoom_id = zoom_id
        self.zoom_join_url = zoom_url
        self.zoom_passcode = zoom_passcode

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


def gather_search_results_by_name(names):
    result = []
    for name in names.split(';'):
        result.extend(database.search_by_name(name.strip()))
    return list(set(result))


def gather_search_results_by_date(year, month, day):
    return database.events_on_date(datetime.datetime(int(year), int(month), int(day)).date())


def sort_results_by_name(events: list):
    return sorted(events, key=lambda event: event.invitee)


def sort_results_by_date(events: list):
    return sorted(events, key=lambda event: event.date)


def sort_results_by_time(events: list):
    return sorted(events, key=lambda event: event.time)


def sort_results(events, sort_by):
    if sort_by == 'Name':
        return sort_results_by_name(events)
    if sort_by == 'Date':
        return sort_results_by_date(events)
    if sort_by == 'Time':
        return sort_results_by_time(events)


def display_results(results: list):
    return [f'{event.invitee}, {event.convert_date()},'
            f' {event.time} ({event.get_standard_time()})' for event in results]


def display_expanded_result(event, values):
    if not event:
        return 'No Result Selected.'
    extended = ''
    if values['-grab_name-']:
        extended += f'Name: {event.invitee}\n'
    if values['-grab_email-']:
        extended += f'Email: {event.invitee_email}\n'
    if values['-grab_timezone-']:
        extended += f'Timezone: {event.invitee_timezone}\n'
    if values['-grab_date-']:
        extended += f'Date: {event.date}\n'
    if values['-grab_time-']:
        extended += f'Time: {event.time}\n'
    if values['-grab_link-']:
        extended += f'Zoom Link: {event.zoom_join_url}\n'
    if values['-grab_passcode-']:
        extended += f'Zoom Passcode: {event.zoom_passcode}\n'
    if values['-grab_id-']:
        extended += f'Zoom ID: {event.zoom_id}\n'
    if values['-grab_notified-']:
        extended += f'Notified: {event.notified}\n'
    return extended


def check_values(values):
    explanation = ''
    if '@' not in values['-email-']:
        explanation += 'Invalid Email.\n'
    if type(values['-year-']) != int:
        explanation += 'Year Must Be An Integer.\n'
    if type(values['-month-']) != int:
        explanation += 'Month Must Be An Integer.\n'
    elif values['-month-'] > 12 or values['-month-'] < 1:
        explanation += 'Month Is Not Valid.\n'
    if type(values['-day-']) != int:
        explanation += 'Day Must Be An Integer.\n'
    elif values['-day-'] > 32 or values['-day-'] < 1:
        explanation += 'Day Is Not Valid.\n'
    if type(values['-hour-']) != int:
        explanation += 'Hour Must Be An Integer.\n'
    elif values['-hour-'] > 12 or values['-hour-'] < 1:
        explanation += 'Hour Is Not Valid.\n'
    if type(values['-minute-']) != int:
        explanation += 'Minute Must Be An Integer.\n'
    elif values['-minute-'] > 60 or values['-minute-'] < 0:
        explanation += 'minute Is Not Valid.\n'
    if values['-zoom_url-'] and 'zoom.us' not in values['-zoom_url-']:
        explanation += 'Zoom Link Is Not Valid.'
    return explanation


def add_new_event():
    timezones = ['Pacific Standard Time', 'Mountain Standard Time', 'Central Standard Time', 'Eastern Standard Time']
    fill_out = [
        [sg.Text('* Name:')],
        [sg.InputText(key='-name-')],
        [sg.Text('* Email:')],
        [sg.InputText(key='-email-')],
        [sg.Text('* Date (YY:MM:DD):')],
        [sg.Spin(list(range(datetime.date.today().year - 10, datetime.date.today().year + 2)),
                 datetime.date.today().year, key='-year-', size=(5, 5)), sg.Text('-'),
         sg.Spin(list(range(1, 13)), datetime.date.today().month, size=(5, 5), key='-month-'), sg.Text('-'),
         sg.Spin(list(range(1, 32)), datetime.date.today().day, size=(5, 5), key='-day-')],
        [sg.Text('* Time (HH:MM am/pm):')],
        [sg.Spin(list(range(1, 13)), 4, size=(5, 5), key='-hour-'), sg.Text(':'),
         sg.Spin(list(range(61)), 0, size=(5, 5), key='-minute-'), sg.Text(':'),
         sg.Spin(['AM', 'PM'], 'PM', size=(5, 5), key='-am_pm-')],
        [sg.Text('Timezone:')],
        [sg.Combo(timezones, 'Pacific Standard Time', key='-timezone-')],
        [sg.Text('Zoom Url (Zoom data will be created if left blank):')],
        [sg.InputText(key='-zoom_url-')],
        [sg.Text('Zoom ID:')],
        [sg.InputText(key='-zoom_id-')],
        [sg.Text('Zoom Passcode:')],
        [sg.InputText(key='-zoom_passcode-')],
        [sg.Text('Notified:')],
        [sg.Combo(['False', 'True'], 'False', key='-notified-')],
        [sg.Text('\n')],
        [sg.Button(' Submit and Save ', key='-submit-'), sg.Button(' Exit ', key='-exit-')]
    ]

    layout = [fill_out]
    window = sg.Window('Add New Events.', layout, resizable=True)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == '-exit-':
            break
        if event == '-submit-':
            response = check_values(values)
            if not response:

                time_ = ''
                if values['-am_pm-'].lower().strip() == 'pm':
                    if int(values['-hour-']) != 12:
                        time_ += str(int(values['-hour-']) + 12)
                    else:
                        time_ += str(int(values['-hour-']))
                else:
                    if int(values['-hour-']) < 10:
                        time_ += '0' + str(int(values['-hour-']))
                    elif int(values['-hour-']) == 12:
                        time_ += '00'
                    else:
                        time_ += str(int(values['-hour-']))
                time_ += ':'
                if int(values['-minute-']) < 10:
                    time_ += '0' + str(int(values['-minute-']))
                else:
                    time_ += str(int(values['-minute-']))
                time_ += ':00'

                date = str(values['-year-']) + '-'
                if int(values['-month-']) < 10:
                    date += '0' + str(int(values['-month-']))
                else:
                    date += str(int(values['-month-']))
                date += '-'
                if int(values['-day-']) < 10:
                    date += '0' + str(int(values['-day-']))
                else:
                    date += str(int(values['-day-']))

                new_event = Event(values['-name-'], values['-email-'], time_, date,
                                  True if values['-notified-'] == 'True' else False, values['-timezone-'],
                                  values['-zoom_url-'], values['-zoom_passcode-'], str(values['-zoom_id-']))
                database.append_event(new_event, zoom, calendar)
                save_db(database)
                sg.Popup('The Event Was Added Successfully!')
            else:
                sg.Popup(response)

    window.close()


def start_window():
    global database
    search_results = sort_results_by_name(gather_search_results_by_name(''))
    extended_view_event = None
    menu = [['Run', 'Core Script']]

    search = [
        [sg.Menu(menu)],
        [sg.Text('Search By Name:')],
        [sg.InputText(key='-name_searched-', enable_events=True), sg.Button(' Search ', key='-search_name-')],
        [sg.Text('')],
        [sg.Text('Search By Date (YY:MM:DD):')],
        [sg.Spin(list(range(datetime.date.today().year - 10, datetime.date.today().year + 2)),
                 datetime.date.today().year, key='-year-', size=(5, 5)), sg.Text('-'),
         sg.Spin(list(range(1, 13)), datetime.date.today().month, size=(5, 5), key='-month-'), sg.Text('-'),
         sg.Spin(list(range(1, 32)), datetime.date.today().day, size=(5, 5), key='-day-'), sg.Text(' '),
         sg.Button(' Search ', key='-search_date-')],
        [sg.Text('\n')],
        [sg.HSeparator()],
        [sg.Text('Sort Search By: '), sg.Combo(['Name', 'Date', 'Time'], 'Name',
                                               enable_events=True, size=(15, 1), key='-sort_by-')],
        [sg.Listbox(display_results(search_results), s=(52, 10), key='-results-', bind_return_key=True)]
    ]

    info = [
        [sg.Frame('Expanded Result Will Show:',
                  layout=[[sg.Checkbox('Name', True, key='-grab_name-', enable_events=True),
                           sg.Checkbox('Email', True, key='-grab_email-', enable_events=True),
                           sg.Checkbox('Timezone', True, key='-grab_timezone-', enable_events=True),
                           sg.Checkbox('Date', True, key='-grab_date-', enable_events=True),
                           sg.Checkbox('Time Start', True, key='-grab_time-', enable_events=True)],
                          [sg.Checkbox('Zoom Link', True, key='-grab_link-', enable_events=True),
                           sg.Checkbox('Zoom Passcode', True, key='-grab_passcode-', enable_events=True),
                           sg.Checkbox('Zoom ID', True, key='-grab_id-', enable_events=True),
                           sg.Checkbox('Notified', True, key='-grab_notified-', enable_events=True)]])],
        [sg.Text('Expanded Result:')],
        [sg.Multiline('', size=(54, 15), key='-expanded_view-')],
        [sg.Button(' Add New Event ', key='-add_event-', size=(22, 1)), sg.Text('  '),
         sg.Button(' Delete Selected Event ', key='-delete-', size=(22, 1))]
    ]

    layout = [[sg.Column(search, vertical_alignment='Top'), sg.VSeparator(),
               sg.Column(info, vertical_alignment='Top')]]

    window = sg.Window('Tutoring Database Viewer.', layout)

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED:
            break
        database = open_db()
        if event == '-search_name-' or event == '-name_searched-':
            search_results = gather_search_results_by_name(values['-name_searched-'])
            window['-results-'].update(display_results(sort_results(search_results, values['-sort_by-'])))
        if event == '-search_date-':
            search_results = gather_search_results_by_date(values['-year-'], values['-month-'], values['-day-'])
            window['-results-'].update(display_results(sort_results(search_results, values['-sort_by-'])))
        if event == '-sort_by-':
            window['-results-'].update(display_results(sort_results(search_results, values['-sort_by-'])))
        if event == '-results-':
            result = values['-results-'][0].split(', ')
            extended_view_event = database.search_by_name_date_time(result[0], result[1], result[2].split(' ')[0])
            window['-expanded_view-'].update(display_expanded_result(extended_view_event, values))
        if 'grab' in event:
            window['-expanded_view-'].update(display_expanded_result(extended_view_event, values))
        if event == '-add_event-':
            window.disable()
            add_new_event()
            window.enable()
            window.bring_to_front()
        if event == '-delete-':
            if not extended_view_event:
                sg.Popup('No Event Has Been Selected.')
                continue
            # Ensuring the event signature is the same as the one on file.
            extended_view_event = database.search_by_name_date_time(extended_view_event.invitee,
                                                                    extended_view_event.date,
                                                                    extended_view_event.time)
            if not database.remove_event(extended_view_event):
                sg.Popup('Count Not Find Event In Database.')
                continue
            for count, event in enumerate(search_results):
                if (event.invitee, event.date, event.time) == (extended_view_event.invitee,
                                                               extended_view_event.date,
                                                               extended_view_event.time):
                    del search_results[count]
            extended_view_event = None
            window['-expanded_view-'].update(display_expanded_result(extended_view_event, values))
            window['-results-'].update(display_results(sort_results(search_results, values['-sort_by-'])))
            save_db(database)
        if event == 'Core Script':
            main.main()
            database = open_db()

    window.close()


def _main():
    global database, zoom, calendar

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

    database = open_db()

    # sg.theme('DarkPurple4')
    sg.theme('DarkBlack')
    start_window()


if __name__ == '__main__':
    _main()
