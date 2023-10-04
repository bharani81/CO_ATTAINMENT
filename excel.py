import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
import os
wb = load_workbook('14CS350_Data Structures and Algorithms.xlsx')

test1 =wb['Test 1']
print(test1.max_row)

test2 = wb['Test 2']
print(test2)

test3 =wb['Test 3']
print(test3)

def create_workbook():
    Wb_test = Workbook()
    sheet = Wb_test.active
    sheet.title = "overall"
    Wb_test.save('new_summary_data.xlsx')


create_workbook()


def create_tot(test):
    i=1
    dict={}
    row = test[3][1:]
    for cos in row:
        if cos.value ==  None:
            return dict
        if cos.value not in dict:
            dict[cos.value]=0
        dict[cos.value]+= int(test[4][i].value) if test[4][i].value != None else 0
        i=i+1

def create_sum(test,r1):
    dict={}
    i=1
    row = test[3][1:]
    for cos in row:
        if cos.value ==  None:
            return dict
        if cos.value not in dict:
            dict[cos.value]=0
        dict[cos.value]+= int(test[r1][i].value) if test[r1][i].value != None else 0
        i=i+1


def foreveryrow(test):
    total_dict = create_tot(test)
    temp_list = list(total_dict.keys())
    choice_co = temp_list[-1]
    total_dict[choice_co] = total_dict[choice_co]/2
    
    new_sheet = load_workbook('new_summary_data.xlsx')
    new_wb = new_sheet['overall']
    new_wb.insert_rows(test.max_row)
    new_sheet.save('new_summary_data.xlsx')
    i=0
    for regno in range(5,test.max_row):
        new_wb.cell(row= regno-2, column= 1, value=test[regno][0].value)
    new_sheet.save('new_summary_data.xlsx')

    c=2
    for keys in total_dict:
        new_wb.cell(row=1,column=c,value=keys)
        new_wb.cell(row=2,column=c,value=total_dict[keys])
        new_wb.cell(row=1,column=c+len(total_dict),value= str(keys)+'%')
        c= c + 1
    new_sheet.save('new_summary_data.xlsx')

    for rows in range(5,test.max_row):
        val_dict = create_sum(test,rows)
        c=2
        for keys in val_dict:
            new_wb.cell(row=rows-2,column=c,value=val_dict[keys])  
            new_wb.cell(row=rows-2,column=c+len(val_dict),value=round(val_dict[keys]/total_dict[keys],2)*100) 
            c=c+1 
    new_sheet.save('new_summary_data.xlsx')


foreveryrow(test3)