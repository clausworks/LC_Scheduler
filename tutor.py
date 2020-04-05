from dataclasses import dataclass
from datetime import time

@dataclass
class Tutor:
    fname: str
    lname: str
    shifts: list
    subjects: list = None

    def __str__(self):
        return '{} {}'.format(self.fname, self.lname)
        # return "{}: {}/{}".format(self.name,self.shifts,self.subjects)


@dataclass
class Shift:
    day: str
    name: str
    inWriting: bool
    # Strings for Days of the Week
    MONDAY = 'Mon'
    TUESDAY = 'Tue'
    WEDNESDAY = 'Wed'
    THURSDAY = 'Thu'
    FRIDAY = 'Fri'
    DAYS = [MONDAY, TUESDAY,
            WEDNESDAY, THURSDAY, FRIDAY]

    def __str__(self):
        return '{} {}'.format(
            self.name[-1],
            self.day)

# @dataclass
# class Shift:
#     start_time: time
#     end_time: time
#     day: str
#
#     def set_start_time(string):
#         pass
#
#     def __str__(self):
#         return "{}-{}".format(self.start_time, self.end_time)

@dataclass
class Subject:
    name: str
    number: str

    def __str__(self):
        return '{} {}'.format(self.name, self.number)
