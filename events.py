import pandas as pd
from icalendar import Calendar, Event, Alarm
from datetime import datetime, timedelta
from pytz import UTC , timezone
import pytz
import yaml

# # Read Excel data
# df = pd.read_excel('events.xlsx') 

ist = timezone('Asia/Kolkata')

# Load the YAML file
with open('events.yaml', 'r') as f:
    data = yaml.safe_load(f)

# events = ["2023-kumba-invite-all-events.ics","2023-kumba-invite-main.ics"]

events = ["main","all"]


def generate_calendar(type):
    # Create a calendar
    cal = Calendar()

    cal.add('X-WR-CALNAME', 'கும்பாபிஷேகம் - காடூர் அருள்மிகு படைகாத்த ஐயனார் கோவில்')

    # for index, row in df.iterrows():
    for event_data in data['events']:

        # startdate_ist = datetime.strptime(event_data['startdate'], '%Y-%m-%d %H:%M:%S')
        # startdate_ist = ist.localize(startdate_ist)

        # # startDate = pd.to_datetime(df['startdate']).dt.tz_localize('Asia/Kolkata')
        # # startDateUTC = startDate.dt.tz_convert(pytz.UTC)

        # # endDate = pd.to_datetime(df['enddate']).dt.tz_localize('Asia/Kolkata')
        # # endDateUTC = endDate.dt.tz_convert(pytz.UTC)





        event = Event()
        event.add('uid', event_data['uid'])
        event.add('summary', event_data['summary'])

        startdate_ist = ist.localize(event_data['startdate'])
        startdate_str = startdate_ist.strftime("%Y%m%dT%H%M%S")

        enddate_ist = ist.localize(event_data['enddate'])
        enddate_str = enddate_ist.strftime("%Y%m%dT%H%M%S")

        event.add('dtstart', startdate_ist)
        event.add('dtend', enddate_ist)
        event.add('description', event_data['description'])
        
        event.add('dtstamp', datetime.now(ist))
        event.add('last-modified', datetime.now(ist))
        event.add('created', datetime.now(ist))
        event.add('sequence', event_data['sequence'])
        event.add('organizer', 'mailto:invite@kovil.org')
    
        event.add('status', "CONFIRMED")
        event.add('class', "PUBLIC")
        event.add('url', 'https://kovil.org')
        
        
        if (event_data['temple-no'] == 1):
            event.add('geo', (11.2607868,79.1028115))
            event.add('location', 'https://goo.gl/maps/eg6E9VAZC6NXpwbR6?coh=178572&entry=tt')
        elif (event_data['temple-no'] == 2):
            event.add('geo', (11.2617603,79.102598))
            event.add('location', 'https://goo.gl/maps/1e211fZgLhBAgmceA?coh=178572&entry=tt')
        elif (event_data['temple-no'] == 3):
            event.add('geo', (11.2627532,79.1030085))
            event.add('location', 'https://goo.gl/maps/LDQifJKwSmR7bKez6?coh=178572&entry=tt')
        
        categories = event_data['categories'].split(',')
        event.add('categories', categories)

        event.add('priority', event_data['priority'])

        if (event_data['priority'] > 3):
            event.add('color', "#00FF00")
            event.add('transp', "OPAQUE")
        else:
            event.add('color', "#000000")
            event.add('transp', "TRANSPARENT")

        if (event_data['reminder-15m'] == 'Y'):
            event.add('valarm', timedelta(minutes=15))
            # alarm = Alarm()
            # alarm.add('action', 'DISPLAY')
            # alarm.add('description', event_data['summary'])
            # alarm.add('trigger', timedelta(minutes=-15))
            # event.add_component(alarm)

        if(event_data['reminder-1H'] == 'Y'):
            # alarm = Alarm()
            # alarm.add('action', 'DISPLAY')
            # alarm.add('description', event_data['summary'])
            # alarm.add('trigger', timedelta(hours=-1))
            event.add('valarm', timedelta(hours=1))  
            # event.add_component(alarm)

        if(event_data['reminder-2H'] == 'Y'):
            # alarm = Alarm()
            # alarm.add('action', 'DISPLAY')
            # alarm.add('description', event_data['summary'])
            # alarm.add('trigger', timedelta(hours=-2))
            event.add('valarm', timedelta(hours=2))
            # event.add_component(alarm)

        if(event_data['reminder-12H'] == 'Y'):
            # alarm = Alarm()
            # alarm.add('action', 'DISPLAY')
            # alarm.add('description', event_data['summary'])
            # alarm.add('trigger', timedelta(hours=-12))
            event.add('valarm', timedelta(hours=12))
            # event.add_component(alarm)

        if (event_data['priority'] == 1 and type == 'main') or type == 'all':
            cal.add_component(event)
 
    filename = "2023-kumba-invite-all-events.ics"
    if (type == "main"):
        filename = "2023-kumba-invite-main-events.ics"

    # Write to file
    with open(filename, 'wb') as f:
        f.write(cal.to_ical())

for event in events:
    generate_calendar(event)