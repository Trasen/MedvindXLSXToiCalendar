import argparse
import os
import sys
from datetime import timedelta

from icalendar import Calendar, Event
import pandas as pd
import datetime

from numpy import ndarray

from Entry import Entry


def main(employee_name, folder_path):
    entries: list[Entry] = []
    walk = os.walk(folder_path)

    files: list[str] | None = None
    for (path, dir, files) in walk:
        files = files

    if files is None:
        sys.stderr.write(f"No directory found on path: '{folder_path}'")
        exit()

    if len(files) < 1:
        sys.stderr.write(f"No files found in folder '{folder_path}'")
        exit()

    for excel in files:

        try:
            file = pd.read_excel(f'{folder_path}/{excel}')
        except:
            sys.stderr.write(f"file:{excel} was not readable, skipping")
            continue

        numpyArray: ndarray = file.to_numpy()
        numpyArray = numpyArray[1:len(numpyArray)]  # Remove the top row from the document

        cell: ndarray
        for cell in numpyArray:

            text_content: str
            for text_content in cell:
                text_content_split = text_content.split("\n")
                text_content_split = list(filter(None, text_content_split))  # Remove empty new lines

                if len(text_content_split) > 1:
                    day_date = datetime.datetime.strptime(text_content_split[0], "%Y-%m-%d")

                    if text_content_split[1] == 'Se':  # if cell contains {date} Se it means vacation
                        entries.append(Entry(day=day_date, start=None, end=None, vacation=True))
                        continue

                    if text_content_split[1] == 'BA':  # if cell contains {date} BA it means parental leave
                        entries.append(Entry(day=day_date, start=None, end=None, parentalLeave=True))
                        continue

                    startEnd = text_content_split[1].split('-')

                    timeOff = startEnd[1].split("   ")

                    if (len(timeOff) > 1):
                        handle_scheduled_vacation_and_parental_leave(day_date, entries, timeOff)
                        continue

                    start = [0]
                    end = startEnd[1]

                    entries.append(Entry(day_date, start, end))

            entries.sort(key=lambda x: x.day)
            writeIcs(entries, employee_name)


def handle_scheduled_vacation_and_parental_leave(day_date, entries, timeOff):
    if (timeOff[1] == 'Se'):  # if vacation on previously scheduled day
        entries.append(Entry(day=day_date, start=None, end=None, vacation=True))
    if (timeOff[1] == 'BA'):  # if parental leave on previously scheduled day
        entries.append(Entry(day=day_date, start=None, end=None, parentalLeave=True))


def writeIcs(entries: list[Entry], employee_name):
    icsPath = "./calendar.ics"

    if os.path.exists(icsPath):  # calendar.ics is supposed to be immutable
        os.remove(icsPath)

    with open(icsPath, 'wb') as file:
        calendar = Calendar()
        calendar["X-WR-CALNAME"] = f"{employee_name} Schema"
        calendar.add('prodid', '-//My calendar product//example.com//')
        calendar.add('version', "2.0")
        for entry in entries:
            event = Event()
            if entry.vacation:
                event.add('summary', f"{employee_name} Semester")
            elif entry.child_care:
                event.add('summary', f"{employee_name} föräldraledig")
            else:
                event.add('summary', f'{employee_name} Jobb')
            event['uid'] = f'{entry.day.date()}_{employee_name}'
            event.add('dtstamp', entry.day)
            event = handleEventStartAndEnd(entry, event)
            calendar.add_component(event)
        file.write(calendar.to_ical())


def handleEventStartAndEnd(entry, event):
    if entry.start:
        event.add('dtstart', entry.start)
    else:
        event.add('dtstart', entry.day.date())

    if entry.end:
        event.add('dtend', entry.end)
    else:
        event.add('dtend',
                  entry.day + timedelta(days=1))  # timedelta days=1 increments the day by 1, this is for all day events
    return event


parser = argparse.ArgumentParser("Converts Medvind schedule xlsx documents to a ICS file")
parser.add_argument("-n", "--employee_name", help="Name of employee", required=False)
parser.add_argument("-f", "--folder_path", help="Path to folder with schedules", required=False, default="./schema")

args = parser.parse_args()
employee_name = args.employee_name
folder_path = args.folder_path

main(employee_name, folder_path)
