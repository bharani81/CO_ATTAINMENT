import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
import os
wb = load_workbook('14CS350_Data Structures and Algorithms.xlsx')

test1 =wb['Test 1']
print(test1)

test2 = wb['Test 2']
print(test2)

test3 =wb['Test 3']
print(test3)

def find_no_columns(test,row):
    c1=1
    for col in test[row]:
        if col!=None:
            c1=c1+1
        else:
            break
    return c1

def create_workbook():
    Wb_test = Workbook()
    sheet = Wb_test.active
    sheet.title = "overall"
    Wb_test.save('new_summary_data.xlsx')

test_list =[test1,test2,test3]

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


def foreveryrow(test,max_row):
    total_dict = create_tot(test)
    temp_list = list(total_dict.keys())
    choice_co = temp_list[-1]
    total_dict[choice_co] = total_dict[choice_co]/2
    
    new_sheet = load_workbook('new_summary_data.xlsx')
    new_wb = new_sheet['overall']
    # new_wb.insert_rows(test.max_row)
    new_sheet.save('new_summary_data.xlsx')
    i=0
    for regno in range(5,test.max_row):
        new_wb.cell(row= regno-2, column= 1, value=test[regno][0].value)
    new_sheet.save('new_summary_data.xlsx')

    c2=1
    for col in new_wb[3]:
        if col.value!=None:
            c2=c2+1
        else :
            break
    c=c2

    for keys in total_dict:
        new_wb.cell(row=1,column=c,value=keys)
        new_wb.cell(row=2,column=c,value=total_dict[keys])
        new_wb.cell(row=1,column=c+len(total_dict),value= str(keys)+'%')
        c= c + 1
    new_sheet.save('new_summary_data.xlsx')

    c1=1
    for col in new_wb[3]:
        if col.value!=None:
            c1=c1+1
        else :
            break

    for rows in range(5,max_row):
        val_dict = create_sum(test,rows)
        c=c1
        for keys in val_dict:
            new_wb.cell(row=rows-2,column=c,value=val_dict[keys])  
            new_wb.cell(row=rows-2,column=c+len(val_dict),value=round(val_dict[keys]/total_dict[keys],2)*100) 
            c=c+1 
    new_sheet.save('new_summary_data.xlsx')
    return new_sheet

def calculate_overall():
    max_row = test1.max_row
    for each_test in test_list:
        foreveryrow(each_test,max_row)

def find_cos(wb,row):
    new_dict={}
    for col in wb[1]:
        co = str(col.value)
        if co.endswith('%'):
            co =co[:-1]
            if co not in new_dict:
                new_dict[co]=[]
            new_dict[co].append(wb[row][col.column-1].value)
    return new_dict

def find_max_pair(percentage_list):
    
    max_sum=0
    
    if(len(percentage_list)==1):
        max_sum = percentage_list[0]
    if(len(percentage_list)==2):
        max_sum = round((percentage_list[0]+percentage_list[1])/2,2)
    if(len(percentage_list)==3):
        for i in range(0,len(percentage_list)):
            for j in range(i+1,len(percentage_list)):
                max_sum = max(round((percentage_list[i] + percentage_list[j])/2,2),max_sum)
    return (max_sum +100)/2

def add_assignment():
    calculate_overall()
    new_sheet = load_workbook('new_summary_data.xlsx')
    new_wb = new_sheet['overall']
    col_val = find_no_columns(new_wb,1)
    for rows in range(3,new_wb.max_row):
        percentage_dict = find_cos(new_wb,rows)
        print(percentage_dict)
        i=1
        for keys in percentage_dict.keys():
            new_wb.cell(row=1,column=col_val+i,value=str(keys))
            new_percentage_list = percentage_dict[keys]
            max_val  = find_max_pair(new_percentage_list)
            new_wb.cell(row=rows,column=col_val + i ,value=max_val)
            i=i+1
    new_sheet.save('new_summary_data.xlsx')
add_assignment()