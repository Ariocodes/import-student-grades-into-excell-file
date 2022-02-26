"""
Author: github.com/Ariocodes
This is a code that imports student grades into an Excell file.
"""

from openpyxl import Workbook

file_base = Workbook()
file = file_base.active
file.title = 'stuGrades'
lis = list(' '.join('BCDEFGHIJKLMNOPQRSTUVWXYZ').split())


def excel_result(dictt: dict):
    num = 2
    subNum = 1
    n = 0
    for subject, allSt in dictt.items():
        file[lis[n] + str(subNum)] = subject
        for name, marks in allSt.items():
            *mid, l = marks
            avg = round(sum(mid, l) / (len(mid) + 1), 2)
            file['A' + str(num)] = name
            if avg < 10:
                file[lis[n] + str(num)] = '{} FAILED'.format(avg)
            else:
                file[lis[n] + str(num)] = avg
            num += 1
            if num - 2 == len(allSt):
                num = 2
        n += 1
    file_base.save('studentResults.xls')


students = {
    'math': {'NAME': ["GRADES_AS_INTEGERS_IN_LIST"],
             'NAME2': ["GRADES_AS_INTEGERS_IN_LIST"],},
    
    'science': {'NAME': ["GRADES_AS_INTEGERS_IN_LIST"],
                'NAME2': ["GRADES_AS_INTEGERS_IN_LIST"],},
    
    'Religion': {'NAME': ["GRADES_AS_INTEGERS_IN_LIST"], 
                 'NAME2': ["GRADES_AS_INTEGERS_IN_LIST"],},
    
    # YOU CAN ADD MORE SUBJECTS AND STUDENTS
}
excel_result(students)
