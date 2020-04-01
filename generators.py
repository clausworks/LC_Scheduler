from utils import get_all_shifts, get_all_subjects

def write_shift_subjects(ws,tutors):
    '''
    Generates a shifts vs. names table in the given worksheet
    using data from the given tutors.
    '''
    SHIFT_COL = 0
    SUBJECT_START_COL = 1  # First subject col

    shifts = get_all_shifts(tutors)
    subjects = get_all_subjects(tutors)

    # Write header column (cell coordinates: row=i,col=j)
    for i,shift in enumerate(shifts):
        ws.cell(row=i+2, column=1).value = shift
    for j,subject in enumerate(subjects):
        ws.cell(row=1, column=j+2).value = subjects[j]

    # Write data
    for tutor in tutors:
        for i,shift in enumerate(shifts):
            for j,subject in enumerate(subjects):
                if shift in tutor.shifts and subject in tutor.subjects:
                    ws.cell(row=i+2, column=j+2).value = 'X'
