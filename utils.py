def parse_csv_list(inputValue):
    '''
    Takes a comma-separated list (any values) and returns the
    individual items as a list (converts to string first)
    '''
    return [s.strip() for s in str(inputValue).split(',')]


def get_all_subjects(tutors):
    '''
    Returns a sorted list of all subjects tutored by the given tutors.
    '''
    subjects = set()
    for tutor in tutors:
        subjects.update(tutor.subjects)
    return sorted(subjects)


def get_all_shifts(tutors):
    '''
    Returns a sorted list of all shifts for the given tutors.
    '''
    shifts = set()
    for tutor in tutors:
        shifts.update(tutor.shifts)
    return sorted(shifts)
