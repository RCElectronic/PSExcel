#TeacherCourses.py
#Author: Andrew Colwell (May 2016) Saint John High School

import xlsxwriter       # create excel spreadsheet
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()

'''
The example requires the user to get data from PowerScheduler.
The report is developed from: Reports, Master Schedule List, Submit
Select and copy all the data below the headings, and save into a simple text file.
The program will create an Excel spreadsheet.
'''

#DataSet: Number.Section, Course Name, Expression, Term, Teacher Name, Teacher Dept., Room, Students, Max Seats
#DataSet: 0             , 1          , 2         , 3   , 4           , 5            , 6   , 7       , 8
#DataSet: BEBUA1200.1, Macrame, 1(A), S1, Martha Stewart, , 335, 16, 22

def string2data(line):
    '''
    This function must match the data from PowerScheduler.
    Dictionary based on teacher name.
    '''
    linelist = line.strip().split('\t')     #split based on tab separation
    for item in linelist:
        # data as a list: [Teacher, Term, Period, Course, Room]
        period = linelist[2]
        dataline = [linelist[4],    # Teacher
                    linelist[3],    # Term
                    period.strip('(A)'),         # Period
                    linelist[1],    # Course
                    linelist[6]]    # Room
    return dataline

def add_data(D,key,datastuff):
    '''
    This function takes a dictionary, checks if key is present, and adds
    datastuff as a list to the value (as a list)
    '''
    keyList = list(D.keys())

    if key in keyList:
        oldValue = D[key]
        newValue = oldValue + [datastuff]

        D[key]=newValue
        return D
    else:
        D[key]=[datastuff]
        return D

        
def data2dictionary(filename):
    '''
    Workhorse function
    '''
    data = open(filename,'r')
    lineRead = 0
    carryOn = 1
    masterDict = {}
    while carryOn == 1:
        nextline = data.readline()
        lineRead += 1      
        if nextline != '':
            datalist = string2data(nextline)

            # data as a list: [Teacher, Term, Period, Course, Room]
            teacherName = datalist[0]
            teacherData = datalist[1:]

            masterDict = add_data(masterDict,teacherName,teacherData)

            
        else:
            carryOn = 0

    return masterDict

def xl_print(xlDataDict):
    
    def preparedata(courselist,row,teacher):
        term = courselist[0]
        period = eval(courselist[1])
        course = courselist[2]
        room = courselist[3]
        block = period * 2


        if term == 'S1' or term == 'Q1' or term == 'Q2':
            existingCourse = xlArrayS1[row][period*2]
            xlArrayS1[row][0]=teacher
            xlArrayS1[row][block-1]=room
            if term =='Q1':
                newCourse = course + '\n' + existingCourse
            else:
                newCourse = existingCourse + '\n' + course
            xlArrayS1[row][period*2]=course
            
        elif term == 'S2' or term == 'Q3' or term == 'Q4':
            existingCourse = xlArrayS1[row][period*2]
            xlArrayS2[row][0]=teacher
            xlArrayS2[row][block-1]=room
            if term =='Q3':
                newCourse = course + '\n' + existingCourse
            else:
                newCourse = existingCourse + '\n' + course
            xlArrayS2[row][period*2]=course
        else:
            xlArrayS1[row][0]=teacher
            xlArrayS1[row][block-1]=room
            xlArrayS1[row][period*2]=course
            xlArrayS2[row][0]=teacher
            xlArrayS2[row][block-1]=room
            xlArrayS2[row][period*2]=course

    keys = list(xlDataDict.keys())
    xlArrayS1 = [['' for x in range(21)] for y in range(len(keys)+5)]
    xlArrayS2 = [['' for x in range(21)] for y in range(len(keys)+5)]
    # data as a list: [Teacher, Term, Period, Course, Room]
    keys.sort()
    rowNumber = 4
    for teacher in keys:
        courses = xlDataDict[teacher]
        for course in courses:
            preparedata(course,rowNumber,teacher)
        rowNumber += 1
    
    #now write s1 and s2 to sheets
    xlfile = file_path.rstrip('.txt')+'.xlsx'
    workbook =xlsxwriter.Workbook(xlfile)
    sem1 = workbook.add_worksheet('Sem1')
    sem2 = workbook.add_worksheet('Sem2')
    sem1.write(0,0,'Semester1 Teacher Schedule')
    sem2.write(0,0,'Semester2 Teacher Schedule')
    heading=['Teacher',
             'Room',
             'A-Block',
             'Room',
             'B-Block',
             'Room',
             'C-Block',
             'Room',
             'D-Block',
             'Room',
             'E-Block',
             '',
             'Alt1',
             '',
             'Alt2']
    for i in range(15):
        sem1.write(2,i,heading[i])
        sem2.write(2,i,heading[i])
    
    for row in range(len(xlArrayS1)):
        for col in range(len(xlArrayS1[row])):
            sem1.write(row,col,xlArrayS1[row][col])

    for row in range(len(xlArrayS2)):
        for col in range(len(xlArrayS2[row])):
            sem2.write(row,col,xlArrayS2[row][col])

    colWidth = [20,6.86,23.14,6.86,23.14,6.86,23.14,6.86,23.14,6.86,23.14]
    for i in range(len(colWidth)):
        sem1.set_column(i,i,colWidth[i])
        sem2.set_column(i,i,colWidth[i])
    
    '''
    for row in range(len(xlArrayS1)):
        for col in range(len(xlArrayS1[row])):
            print('Row:',row,' Col:',col,sep='',end='')
        print('\n')
    '''
    
    workbook.close()


#main program           
    
dataDict = data2dictionary(file_path)

xl_print(dataDict)

'''
teachers = list(dataDict.keys())
teachers.sort()
for teacher in teachers:
   print(teacher)
    courses = dataDict[teacher]
    for course in courses:
        print(course)
'''


    

