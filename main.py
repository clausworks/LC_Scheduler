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
    # print(name_shifts)
    # print(name_subjects)
    # TODO: Check tutor names against each other for consistency.
    # Generate tutor objects from data dictionaries.
    tutors = []
    for name in name_shifts:
        tutor = Tutor(name=name,
                      shifts=name_shifts[name],
                      subjects=name_subjects[name])
        tutors.append(tutor)

    if 'Shift-Subjects' not in wb: wb.create_sheet('Shift-Subjects')
    write_shift_subjects(wb['Shift-Subjects'], tutors)
    print('Shift-Subjects sheet generated')

    wb.save(wb_filename)
    print('Done.')

    # Sheet: Generated2
    # for subj in all_subjects:
    #     i = all_subjects.index(subj) + 1
    #     ws_generated2[get_column_letter(i+1)+'1'].value = subj
    #     ws_generated2[get_column_letter(i+1)+'1'].font = Font(bold=True)
    # for row in ws_schedule.iter_rows():
    #     ws_generated2['A'+str(row[0].row+1)] = row[0].value
    # for row in ws_schedule.iter_rows():
    #     tutors = parse_csv_list(row[1].value)
    #     cur_subjects = set()
    #     for t in tutors:
    #         cur_subjects.update(tutor_subjects[t])
    #     for i in range(len(all_subjects)):
    #         coord = get_column_letter(i+2) + str(row[0].row + 1)  # coordinate of cell to place 'X' in
    #         if all_subjects[i] in cur_subjects: ws_generated2[coord].value = 'X'
    # print('Sheet 2 generated.')



if __name__ == '__main__':
    main()
