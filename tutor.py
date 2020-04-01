from dataclasses import dataclass

@dataclass
class Tutor:
    name: str
    shifts: list
    subjects: list

    def __str__(self):
        return "{}: {}/{}".format(self.name,self.shifts,self.subjects)
