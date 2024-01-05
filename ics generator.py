from icalendar import Calendar, Event, Alarm
from datetime import datetime, timedelta
import PySimpleGUI as sg
import pandas as pd


def create_ics_event(calendar, start_datetime, end_datetime, subject, location='',
                     priority=0, private=False, sensitivity='Public', show_time_as='Busy',
                     reminder_minutes_before=15):
    # Create a new event
    event = Event()

    # Set event properties
    event.add('summary', subject)
    event.add('location', location)
    event.add('priority', priority)

    # Set time information
    event.add('dtstart', start_datetime)
    event.add('dtend', end_datetime)

    # Set additional properties
    event.add('private', private)
    event.add('sensitivity', sensitivity)
    event.add('showas', show_time_as)

    # Add a default reminder
    reminder = Alarm()
    reminder.add('action', 'DISPLAY')
    reminder.add('description', f'Reminder: {subject}')
    reminder.add('trigger', timedelta(minutes=-reminder_minutes_before))
    event.add_component(reminder)

    # Add the event to the provided calendar
    calendar.add_component(event)


def create_empty_calendar():
    # Create a new calendar
    cal = Calendar()

    # Set calendar properties
    cal.add('prodid', '-//My Calendar//example.com//')
    cal.add('version', '2.0')

    return cal


def GUI_code():
    layout = [
        [sg.Text('Please fill out the following fields:')],
    ]

    days = ['M', 'T']
    times = ['1', '2', '3', '4', '5', '6', '7']

    for day in days:
        for time in times:
            class_key = f'Class_{day}{time}'
            location_key = f'Location_{day}{time}'

            layout.append([
                sg.Text(f'{day}{time}', size=(15, 1)),
                sg.Text('Class', size=(5, 1)),
                sg.InputText(key=class_key, size=(15, 1)),
                sg.Text('Location', size=(8, 1)),
                sg.InputText(key=location_key, size=(15, 1))
            ])

    layout.append([sg.Submit(), sg.Button('Clear'), sg.Exit()])

    window = sg.Window('Data Entry Form', layout)

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Exit':
            break

        if event == 'Clear':
            for key in values:
                window[key]('')
            continue

        if event == 'Submit':
            m_day_classes = [values.get(f'Class_{day}{time}', '') for day in ['M'] for time in times]
            t_day_classes = [values.get(f'Class_{day}{time}', '') for day in ['T'] for time in times]
            m_day_locations = [values.get(f'Location_{day}{time}', '') for day in ['M'] for time in times]
            t_day_locations = [values.get(f'Location_{day}{time}', '') for day in ['T'] for time in times]

            window.close()
            return m_day_classes, t_day_classes, m_day_locations, t_day_locations

    window.close()


def categorize_subjects(csv_file_path):
    # Read the CSV file
    df = pd.read_csv(csv_file_path)

    # Filter relevant columns
    df = df[['Subject', 'Start Date']]

    # Handle NaN values in 'Subject' column
    df['Subject'] = df['Subject'].apply(lambda x: 'Other' if pd.isna(x) else x)

    # Create a new column 'Category' based on the 'Subject' values
    df['Category'] = df['Subject'].apply(lambda x: categorize_subject(x))

    # Remove rows with 'Other' category
    df = df[df['Category'] != 'Other']

    return df[['Subject', 'Start Date', 'Category']]

def categorize_subject(subject):
    # Categorize subjects based on the specified criteria
    if 'M' in subject and any(char.isdigit() for char in subject) and '- ECDT SSOC' not in subject:
        return 'M'
    elif 'T' in subject and any(char.isdigit() for char in subject) and '- ECDT SSOC' not in subject:
        return 'T'
    elif 'M' in subject and any(char.isdigit() for char in subject) and '- ECDT SSOC' in subject:
        return 'SM'
    elif 'T' in subject and any(char.isdigit() for char in subject) and '- ECDT SSOC' in subject:
        return 'ST'
    else:
        return 'Other'


def generate_ics_events(result_df, result_m_day_classes, result_m_day_locations, result_t_day_classes,
                        result_t_day_locations, cal):
    for index, row in result_df.iterrows():
        category = row['Category']
        entry_date = datetime.strptime(row['Start Date'], '%m/%d/%Y')

        if category in ['M', 'T']:
            period_start_times = period_to_start_time
            period_end_times = period_to_end_time
            classes = result_m_day_classes if category == 'M' else result_t_day_classes
            locations = result_m_day_locations if category == 'M' else result_t_day_locations
        elif category in ['SM', 'ST']:
            period_start_times = period_to_start_time_SSOC
            period_end_times = period_to_end_time_SSOC
            classes = result_m_day_classes if category == 'SM' else result_t_day_classes
            locations = result_m_day_locations if category == 'SM' else result_t_day_locations
        else:
            continue

        for period, class_entry in enumerate(classes):
            if class_entry:
                start_time = datetime.combine(entry_date,
                                              datetime.strptime(period_start_times[period + 1], '%I:%M:%S %p').time())
                end_time = datetime.combine(entry_date,
                                            datetime.strptime(period_end_times[period + 1], '%I:%M:%S %p').time())

                create_ics_event(cal, start_time, end_time, class_entry, locations[period])


################# main() #####################################
csv_file_path = 'Spring 2024 M-T Day Calendar.CSV'
result_df = categorize_subjects(csv_file_path)
print(result_df)
period_to_start_time = {1: '7:30:00 AM', 2: '8:30:00 AM', 3: '9:30:00 AM', 4: '10:30:00 AM', 5: '12:45:00 PM',
                        6: '1:45:00 PM', 7: '2:45:00 PM'}
period_to_end_time = {1: '8:30:00 AM', 2: '9:30:00 AM', 3: '10:30:00 AM', 4: '11:30:00 AM', 5: '1:45:00 PM',
                      6: '2:45:00 PM', 7: '3:45:00 PM'}
period_to_start_time_SSOC = {1: '7:30:00 AM',
                             2: '8:27:00 AM',
                             3: '9:24:00 AM',
                             4: '10:21:00 AM',
                             5: '1:22:00 PM',
                             6: '2:19:00 PM',
                             7: '3:16:00 PM'}
period_to_end_time_SSOC = {1: '8:20:00 AM', 2: '9:17:00 AM', 3: '10:14:00 AM', 4: '11:11:00 AM', 5: '2:12:00 PM',
                           6: '3:09:00 PM', 7: '4:06:00 PM'}

result_m_day_classes, result_t_day_classes, result_m_day_locations, result_t_day_locations = GUI_code()
print("M-day classes:", result_m_day_classes)
print("T-day classes:", result_t_day_classes)
print("M-day locations:", result_m_day_locations)
print("T-day locations:", result_t_day_locations)

cal = create_empty_calendar()  # Assume create_empty_calendar function is defined as before
generate_ics_events(result_df, result_m_day_classes, result_m_day_locations, result_t_day_classes,
                    result_t_day_locations, cal)
# Now 'cal' contains the generated events, and you can save it to a file or use it as needed.
with open('calendar_with_event.ics', 'wb') as f:
    f.write(cal.to_ical())

