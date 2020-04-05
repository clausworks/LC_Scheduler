from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from tutor import Tutor
from collectors import read_name_shifts, read_name_subjects
from generators import write_shift_names, write_shift_subjects
from utils import get_writing_tutors, get_nonwriting_tutors, get_all_subjects_dict

def main():
    wb_filename = 'schedule.xlsx'
    wb = load_workbook(filename=wb_filename, read_only=True)

    # Read data into dictionaries.
    name_shifts = read_name_shifts(wb['Tutor Schedule'])
    name_subjects = read_name_subjects(wb['Subjects'])

    for name in name_subjects:
        if name not in name_shifts:
            print(name,"'s schedule can't be found.", sep='')
    # TODO: Check tutor names both ways for consistency.

    # Generate tutor objects from data dictionaries.
    tutors = []
    for name in name_shifts:
        tutor = Tutor(lname=name[0],
                      fname=name[1],
                      shifts=name_shifts[name])
        tutors.append(tutor)

    # writing_tutors = get_writing_tutors(tutors)
    nonwriting_tutors = get_nonwriting_tutors(tutors)
    for tutor in nonwriting_tutors:
        cur_name = (tutor.lname,tutor.fname)
        if cur_name in name_subjects:
            tutor.subjects = name_subjects[cur_name]
        else:
            pass  # FIXME: add further error checking?

    # get_all_subjects_dict(nonwriting_tutors)

    # print('\nWRITING TUTORS')
    # for tutor in w_tutors:
    #     print(tutor.fname, tutor.lname)
    # print('\nNON-WRITING TUTORS')
    # for tutor in nw_tutors:
    #     print(tutor.fname, tutor.lname)
    # print('\nUNCLASSIFIED TUTORS')
    # for tutor in tutors:
    #     if tutor not in w_tutors and tutor not in nw_tutors:
    #         print(tutor.fname, tutor.lname)


    # Write to workbook.
    new_wb = Workbook()

    ws_gtma_name = 'GT & Math (by Tutor)'
    new_wb.active.title = ws_gtma_name  # Rename default new sheet.
    write_shift_names(new_wb[ws_gtma_name], nonwriting_tutors)

    ws_gtma_subj_name = 'GT & Math (by Subject)'
    new_wb.create_sheet(ws_gtma_subj_name)
    write_shift_subjects(new_wb[ws_gtma_subj_name], nonwriting_tutors)

    new_wb.save('GENERATED.xlsx')
    print('Done.')

if __name__ == '__main__':
    main()
