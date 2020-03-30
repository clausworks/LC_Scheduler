from openpyxl import load_workbook

def parse(inputValue):
    '''
    Takes a comma-separated list (any value) and returns the
    individual items as a list (converts to string first)
    '''
    return [s.strip() for s in str(inputValue).split(',')]

def main():
    wb_filename = 'Book1.xlsx'
    wb = load_workbook(filename=wb_filename)
    ws_schedule = wb['Schedule']
    ws_subjects = wb['Subjects']
    
    tutorSubjects = {row[0].value:parse(row[1].value) for row in ws_subjects.rows}
    
    ws_schedule.insert_cols(2)
    wb.save(wb_filename)
    for row in ws_schedule.iter_rows():
        tutors = parse(row[2].value)
        outputSet = set()
        for t in tutors:
            outputSet.update(tutorSubjects[t])
        row[1].value = ', '.join(sorted(outputSet))

    wb.save(wb_filename)


if __name__ == '__main__':
    main()


