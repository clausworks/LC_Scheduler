from dataclasses import dataclass
from datetime import time

@dataclass
class Tutor:
    fname: str
    lname: str
    shifts: list
    subjects: list
    shifts: list

    def __str__(self):
        return "{}: {}/{}".format(self.name,self.shifts,self.subjects)


@dataclass
class Shift:
    start_time: time
    end_time: time
    day: str

    def set_start_time(string):
        pass

    def __str__(self):
        return "{}-{}".format(self.start_time, self.end_time)
