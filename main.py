from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from tutor import Tutor
from collectors import read_name_shifts, read_name_subjects
from generators import write_shift_subjects

def main():
    wb_filename = 'schedule.xlsx'
    wb = load_workbook(filename=wb_filename)

    # Read data into dictionaries.
    name_shifts = read_name_shifts(wb['Name-Shifts'])
    name_subjects = read_name_subjects(wb['Name-Subjects'])
    # TODO: Check tutor names against each other for consistency.

    # Generate tutor objects from data dictionaries.
    tutors = []
    for name in name_shifts:
        tutor = Tutor(name=name,
                      shifts=name_shifts[name],
                      subjects=name_subjects[name])
        tutors.append(tutor)

    # Write to workbook.
    if 'Shift-Subjects' not in wb: wb.create_sheet('Shift-Subjects')
    write_shift_subjects(wb['Shift-Subjects'], tutors)
    print('Shift-Subjects sheet generated')

    wb.save(wb_filename)
    print('Done.')

if __name__ == '__main__':
    main()
