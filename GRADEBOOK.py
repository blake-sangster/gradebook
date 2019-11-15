#importing modules to help read and write in excel
import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell
import xlwt 
from xlwt import Workbook 

#text files to store user info
class_list = open('class_list.txt','w+')
grade_list = open('grade_list.txt','w+')
admins = {'blake'  : 'abc123',' ' : ' '}


#calculates a class grade when giving 5 test score
def whatgrade():
    print('Enter the grades of tests taken, followed by desired upcoming test scores:  ')
    grade1 = input();
    if grade1 == 'x':
        exit()   
    else:
        grade1 = int(grade1)
        grade2 = int(input())
        grade3 = int(input())
        grade4 = int(input())
        grade5 = int(input())

        sum = grade1 + grade2 + grade3 + grade4 + grade5
        average = sum/5

        if average >= 90:
            print('your grade is an A')
            print()
            print('Course Average: ', average)
            print()
            main()
        elif average >= 80:
            print('your grade is a B')
            print()
            print('Course Average: ', average)
            print()
            main()
        elif average >= 70:
            print('your grade is a C')
            print()
            print('Course Average: ', average)
            print()
            main()
        elif average == 69:
            print()
            print('nice.')
            print()
            print('your grade is a D')
            print()
            print('course average: ', average)
            main()
        elif average >= 60:
            print('your grade is a D')
            print()
            print('Course Average: ', average)
            print()
            main()
        elif average >= 59:
            print('your grade is a F')
            print()
            print('Course Average: ', average)
            print()
            main()
        else:
            ('enter a vaild grade')



#main function to add class to text file       
def add_class(classadd):
    class_input = input('Enter class: ')
    count = 0

    while class_input:
        count += 1
        classadd.write(str(count) + ': ' + class_input + '\n')
        class_input = input('Enter class: ')

    return count

#seeks the begging of the class text file and reads what user has entered
def add_class2():
    add_class(class_list)

    class_list.seek(0)
    class_text = class_list.read()

    print()
    print(class_text)
    main()
#reads excel file. .head() is usedto return the first rows in excel file 
def excel():
    df = pd.read_excel('gradebook.xls')
    df.head()
    print(df.head(13))
    main()
#main file to enter grades to text file
def add_grade(gradeadd):
    grade_input = input('Enter grade: ')
    count = 0

    while grade_input:
        count += 1
        gradeadd.write(str(count) + ': ' + grade_input + '\n')
        grade_input = input('Enter Grade: ')

    return count
#reads grade text file and displays message
def add_grade2():
    add_grade(grade_list)

    grade_list.seek(0)
    grade_text = grade_list.read()

    print()
    print(grade_text)
    main()
    
#all users will have to login
def login():
    login = input('User: ')
    password = input('Password: ')

    if login in admins:
        if admins[login] == password:
            print("Welcome " + login,'!')
            print()
            main()
    else:
        invalid_user()
#displays invaild user
def invalid_user():
    print('Invalid user')
    login()

#function to write to excel
def add_excel():
    print('Enter the first grade below, followed by the second and so on. ')
    grade1 = int(input('Math grade 1: '))
    grade2 = int(input('Math grade 2: '))
    grade3 = int(input('Math grade 3: '))
    grade4 = int(input('Science grade 1: '))
    grade5 = int(input('Science grade 2: '))
    grade6 = int(input('Science grade 3: '))
    grade7 = int(input('PE grade 1: '))
    grade8 = int(input('PE grade 2: '))
    grade9 = int(input('PE grade 3: '))
    grade10 = int(input('History grade 1: '))
    grade11 = int(input('History grade 2: '))
    grade12 = int(input('History grade 3: '))

    sum = grade1 + grade2 + grade3
    average1 = sum/3

    sum = grade4 + grade5 + grade6
    average2 = sum/3



        
    # Workbook is created 
    wb = Workbook() 
      
    # add_sheet is used to create sheet. *NOT COMPLETE*
    sheet1 = wb.add_sheet('Sheet 1') 
      
    sheet1.write(1, 0, 'Math') 
    sheet1.write(2, 0, 'Science') 
    sheet1.write(3, 0, 'PE') 
    sheet1.write(4, 0, 'History') 
    sheet1.write(5, 0, 'Coding') 
    sheet1.write(0, 1, 'Test 1') 
    sheet1.write(0, 2, 'Test 2') 
    sheet1.write(0, 3, 'Test 3') 
    sheet1.write(0, 4, 'Average') 
    sheet1.write(1, 1, grade1)
    sheet1.write(1, 2, grade2)
    sheet1.write(1, 3, grade3)
    sheet1.write(1, 4, average1)
    sheet1.write(2, 1, grade4)
    sheet1.write(2, 2, grade5)
    sheet1.write(2, 3, grade6)
    sheet1.write(2, 4, average2)

      
    wb.save('gradebook.xls')
    main()




#reads and displays entered class input
def seeclasses():
    class_list.seek(0)
    class_text = class_list.read()

    print()
    print(class_text)
    main()
def exitnow():
    print('GOODBYE','\nFRIEND')
    print()
    print('By: Blake Sangster')
def viewgrades():
    grade_list.seek(0)
    grade_text = grade_list.read()

    print()
    print(grade_text)
    main()
#main screen that the user sees after login 
def main():
    print('''
                        WELCOME TO THE GRADE TARCKER!

                        (1) Whats my desired grade?
                        (2) Add class scedule
                        (3) View Class scedule
                        (4) Add grades
                        (5) View entered grades
                        (6) Excel gradebook
                        (7) Add excel
                        (8) EXIT
    ''')

    action = input('''
                        What would you like to do? ''')

    if action == '1':
        whatgrade()
    elif action == '2':
        add_class2()
    elif action == '3':
        seeclasses()
    elif action == '4':
        add_grade2()
    elif action == '5':
        viewgrades()
    elif action == '6':
        excel()
    elif action == '7':
        add_excel()
    elif action == '8':
        exitnow()

login()
    

