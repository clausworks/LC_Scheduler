from utils import parse_csv_list
import re
from tutor import Shift, Subject

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

    This function returns a dict of {(lname,fname):[shifts]} (with format
    tuple:list).
    '''

    data = dict()
    cur_tutor_name = ""  # needs to be in scope for multiple loops
    for row_num,row in enumerate(ws.iter_rows(values_only=True)):
        string0 = str(row[0]).lower()
        # Name cells can be in any of the following formats:
        #   - "Brunet, Nicholas (ST0045)"
        #   - "Brunet, Nicholas (ST0045/ST0040)"
        #   - "Brunet, Nicholas (ST00XX)"
        if '(st00' in string0:
            numIndex = string0.find('(st00')
            name_string = str(row[0])[:numIndex]  # Strips off ST number.
            names = name_string.split(',')
            names = [name.strip() for name in names]
            # print('Tutor:',names[1], names[0])
            cur_tutor_name = (names[0], names[1])  # last, first
            data[cur_tutor_name] = []
        # TODO: check for possible undetected name (no ST number)
        # if re.search(r'[A-Za-z\s]+,\s[A-Za-z\s]+', string0):
        #     print('Possible undetected name:', string0)

        # Shift cells begin with a row that includes "Drop-in", e.g:
        #   - "Drop-in Math Shift 2"
        #   - "Drop-in Shift 1"
        elif 'drop' in string0 and 'in' in string0:
            pattern = r'shift [1-3]'  # NOTE: Shifts must be numbered 1-3
            if re.search(pattern, string0):
                span = re.search(pattern, string0).span()
                # Get shift number.
                shift_num = string0[span[0]:span[1]][-1]
                # print(cur_tutor_name)
                # print(row[:10],end='\n\n')
                inWriting = True if 'writing' in string0 else False
                duplicate_shifts = []
                if row[1] and row[2]:
                    new_shift = Shift(Shift.MONDAY,'Shift '+shift_num, inWriting)
                    if new_shift in data[cur_tutor_name]:
                        duplicate_shifts.append(new_shift.name
                            + ' (' + new_shift.day + ')'
                            + 'in row number ' + str(row_num + 1))
                    data[cur_tutor_name].append(new_shift)
                if row[3] and row[4]:
                    new_shift = Shift(Shift.TUESDAY,'Shift '+shift_num, inWriting)
                    if new_shift in data[cur_tutor_name]:
                        duplicate_shifts.append(new_shift.name
                            + ' (' + new_shift.day + ')'
                            + 'in row number ' + str(row_num + 1))
                    data[cur_tutor_name].append(new_shift)
                if row[5] and row[6]:
                    new_shift = Shift(Shift.WEDNESDAY,'Shift '+shift_num, inWriting)
                    if new_shift in data[cur_tutor_name]:
                        duplicate_shifts.append(new_shift.name
                            + ' (' + new_shift.day + ')'
                            + 'in row number ' + str(row_num + 1))
                    data[cur_tutor_name].append(new_shift)
                if row[7] and row[8]:
                    new_shift = Shift(Shift.THURSDAY,'Shift '+shift_num, inWriting)
                    if new_shift in data[cur_tutor_name]:
                        duplicate_shifts.append(new_shift.name
                            + ' (' + new_shift.day + ')'
                            + 'in row number ' + str(row_num + 1))
                    data[cur_tutor_name].append(new_shift)
                if row[9] and row[10]:
                    new_shift = Shift(Shift.FRIDAY,'Shift '+shift_num, inWriting)
                    if new_shift in data[cur_tutor_name]:
                        duplicate_shifts.append(new_shift.name
                            + ' (' + new_shift.day + ')'
                            + 'in row number ' + str(row_num + 1))
                    data[cur_tutor_name].append(new_shift)
                if duplicate_shifts:
                    err_file = open('ERRORS.txt', 'a')
                    print('Duplicate shifts found for ', end='', file=err_file)
                    print(cur_tutor_name[0], ', ', cur_tutor_name[1], sep='', file=err_file)
                    for i,cur_shift in enumerate(duplicate_shifts):
                        print('\t', i+1, ': ',
                            cur_shift,
                            sep='', file=err_file)
                    print('',file=err_file)
                    err_file.close()

    return data


def read_name_subjects(ws):
    '''
    Reads data from the given Excel worksheet in the format names vs. subjects.
    This function returns a dictionary in the form {(lname,fname):[subjects]}

    A row is arranged as follows:
        A: name (last, first) or empty (last non-empty name cell is used)
        B: subject abbreviation (e.g. "BIOL")
        C: comma-separated course numbers
    '''

    data = {}
    last_valid_name = ''
    for row in ws.iter_rows(values_only=True):
        if row[0]:
            name = tuple(parse_csv_list(row[0]))
            last_valid_name = name
        else:
            name = last_valid_name
        subject_area = row[1]
        numbers = parse_csv_list(row[2])
        data[name] = []
        if numbers:
            for cur_number in numbers:
                data[name].append(Subject(subject_area, cur_number))
        else:
            data[name] = [Subject(subject_area, '(All)')]  # No numbers recorded
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
