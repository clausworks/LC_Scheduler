def parse_csv_list(inputValue):
    '''
    Takes a comma-separated list (any values) and returns the
    individual items as a list (converts to string first)
    '''
    values = [s.strip() for s in str(inputValue).split(',')]
    if values == ['None']:
        return None
    else:
        return values

def get_writing_tutors(tutors):
    writing_tutors = []
    for tutor in tutors:
        for shift in tutor.shifts:
            if shift.inWriting:
                writing_tutors.append(tutor)
                break
    return writing_tutors

def get_nonwriting_tutors(tutors):
    nonwriting_tutors = []
    for tutor in tutors:
        if tutor.shifts:
            for shift in tutor.shifts:
                if shift.inWriting:
                    break
            else:
                nonwriting_tutors.append(tutor)
    return nonwriting_tutors


def get_all_subjects_dict(tutors):
    '''
    Returns a sorted list of all subjects tutored by the given tutors.
    Format: {subj_area: numbers_list}
    '''
    all_subjects = dict()
    for tutor in tutors:
        if tutor.subjects:
            for cur_subj in tutor.subjects:
                if cur_subj.name not in all_subjects: all_subjects[cur_subj.name] = set()
                all_subjects[cur_subj.name].add(cur_subj.number)
                # print('Modified:',all_subjects[cur_subj.name])
        else:
            pass  # FIXME: add some type of error cheking
    # Convert sets to lists and sort.
    for k,v in all_subjects.items():
        all_subjects[k] = sorted(list(v))
        print(k,all_subjects[k])
    return all_subjects


def get_all_shifts(tutors):
    '''
    Returns a sorted list of all names (strings) of shifts.
    '''
    all_shifts = set()
    for tutor in tutors:
        for shift in tutor.shifts:
            all_shifts.add(shift.name)
    return sorted(list(all_shifts))
