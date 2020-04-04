from utils import parse_csv_list

def read_name_shifts(ws):
    '''
    Reads data from the given Excel worksheet in the format names vs. shifts.
    The name is gathered from a cell in column A containing "(ST00" where *
    matches any character(s). The cell is in the format "Last, First (ST00*)".

    A shift row is denoted by the text "drop-in" in column A. The format is
    as follows:
        A: indicates subject
        B: start time (HH:MM *M) Monday
        C: end time (HH:MM *M) Monday
        D,E,F,G,H,I,J,K Tues-Fri in like manner

    This function returns a dict of {name:shifts} (with format string:list).
    '''

    # 0-based columns
    NAME = 0
    SHIFT_TYPE = 0
    MON_1 = 1
    MON_2 = 2
    TUE_1 = 3
    TUE_2 = 4
    WED_1 = 5
    WED_2 = 6
    THU_1 = 7
    THU_2 = 8
    FRI_1 = 9
    FRI_2 = 10

    data = {}
    for row in ws.iter_rows(values_only=True):
        # only cells with names have this, e.g. (ST0045)
        index = -1
        try:
            index = str(row[0]).upper().index('(ST00')
        except ValueError:
            continue
        else:
            name_string = str(row[0])[:index]  # strip off number
            names = name_string.split(',')
            names = [name.strip() for name in names]
            print('Tutor:',names[1], names[0])
    # return data


def read_name_subjects(ws):
    '''
    Reads data from the given Excel worksheet in the format names vs. subjects.
    Returns a dictionary in the format <name: subjects>
    '''
    NAME_COL = 0
    SUBJECTS_COL = 1

    data = {}
    for row in ws.iter_rows(values_only=True):
        name = row[NAME_COL]
        subjects = parse_csv_list(row[SUBJECTS_COL])
        data[name] = subjects
    return data


def read_shift_names(ws):
    '''
    Reads data from the given excel worksheet in the format name vs. shifts.
    There are four columns for MTWR. Each cell contains the names of tutors
    who are working that shift on that day.
    '''
    # 0-based columns and rows
    MON_COL = 1
    TUE_COL = 2
    WED_COL = 3
    THU_COL = 4
    SHIFT_1_ROW = 1
    SHIFT_2_ROW = 2
    SHIFT_3_ROW = 3

    #
