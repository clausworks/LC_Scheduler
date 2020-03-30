from openpyxl import load_workbook

def parse(inputValue):
    '''
    Takes a comma-separated list (any value) and returns the
    individual items as a list (converts to string first)
    '''
    return [s.strip() for s in str(inputValue).split(',')]

def main():
    wb = load_workbook(filename='Book1.xlsx')
    ws_schedule = wb['Schedule']
    ws_subjects = wb['Subjects']

    schedule = {}
    i=1
    curVal = ws_schedule['A'+str(i)].value
    while curVal:
        schedule[curVal] = parse(ws_schedule['B'+str(i)].value)

        i+=1
        curVal = ws_schedule['A'+str(i)].value
    for item in schedule.items():
        print(item)

    print('\n')

    subjects = {}
    i=1
    curVal = ws_subjects['A'+str(i)].value
    while curVal:
        subjects[curVal] = parse(ws_subjects['B'+str(i)].value)
        i+=1
        curVal = ws_subjects['A'+str(i)].value
    for item in subjects.items():
        print(item)


if __name__ == '__main__':
    main()


