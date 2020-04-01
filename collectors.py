from utils import parse_csv_list

def read_name_shifts(ws):
    '''
    Reads data from the given Excel worksheet in the format names vs. shifts.
    Shifts are denoted by an 'X' in the appropriate column.
    Returns a dictionary in the format <name: shifts>
    '''
    # 0-based column and row values
    NAME_COL = 0
    FIRST_DATA_ROW = 1;

    shift_names = [cell.value for cell in ws[1]]  # Get first row for shift names
    shift_names = shift_names[1:]  # Chop off empty cell (A1)

    data = {}
    for row in ws.iter_rows(min_row=FIRST_DATA_ROW+1, values_only=True):
        name = row[NAME_COL]
        shifts = [shift_names[i] for i in range(3) if row[i+1]]
        data[name] = shifts
    return data


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
