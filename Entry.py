import datetime
from datetime import timedelta


class Entry:
    def __init__(self, day: datetime, start: str | None, end: str | None, vacation: bool = False,
                 parentalLeave: bool = False):
        self.day = day

        if(start != None):
            (shour, sminute) = start.split(":")
            self.start: datetime = day + timedelta(hours=int(shour), minutes=int(sminute))
        else: self.start = None
        if(end != None):
            (ehour, eminute) = end.split(":")
            self.end: datetime = day + timedelta(hours=int(ehour), minutes=int(eminute))
        else: self.end = day.date()
        self.vacation: bool = vacation
        self.child_care: bool = parentalLeave