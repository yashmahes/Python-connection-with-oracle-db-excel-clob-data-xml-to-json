import xlrd
import csv
import xlsxwriter
import xmltodict
import json
import datetime

import cx_Oracle
#import pandas as pd
import time


import pandas as pd

def convert_xml_to_json(s):
    return json.loads(json.dumps(xmltodict.parse(s)))

data_xls = pd.read_excel('input.xls', 'SubmissionTest (1)', index_col=None)
data_xls.to_csv('your_csv.csv', encoding='utf-8')


file = open('your_csv.csv') 

lines = file.readlines()

arr1 = []
arr2 = []
arr3 = []
arr4 = []
arr5 = []
result = []
exec_time = []

i = 1
while i< len(lines):
    words = (lines[i]).split(",")
    
    a1 = (words[1].strip())
    arr1.append(a1)
    
    a2 = (words[2].strip())
    arr2.append(a2)
    
    a3 = (words[3].strip())
    arr3.append(a3)
    
    a4 = (words[4].strip())
    arr4.append(a4)
    
    a5 = (words[5].strip())
    arr5.append(a5)
    
    my_connection=cx_Oracle.Connection("Clarify/devint$r0ck@159.127.44.236:1521/DEVINT")
    my_cursor=my_connection.cursor()
    myvar = my_cursor.var(cx_Oracle.CURSOR)
    my_cursor.callproc("CLARIFY.AXIS_SEASONALCAMPAIGN_BK.REBATE_CHECK", [a1, a2, a3, a4, a5, myvar])

    
    dane = myvar.getvalue().fetchone()
    
    
    t = str(dane[0])
    #print(t)
    
    exec_time.append(t)
    
    
    res = "no value"
    try:
        res = (dane[1].read())

        res = str(convert_xml_to_json(res))
    except:
        pass
    
    
    result.append(res)
    
    #print(str(dane))
    
    
    
    i+=1
    
    
    

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('output.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 20)
worksheet.set_column('B:B', 20)
worksheet.set_column('C:C', 20)
worksheet.set_column('D:D', 20)
worksheet.set_column('E:E', 20)
worksheet.set_column('F:F', 20)
worksheet.set_column('G:G', 20)
# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})


# Write some simple text.
#worksheet.write('A1', 'Hello')
# Text with formatting.
#worksheet.write('A2', 'World', bold)

# Write some numbers, with row/column notation.
#worksheet.write(2, 0, 123)
#worksheet.write(3, 0, 123.456)

# Insert an image.
# worksheet.insert_image('B5', 'logo.png')




words = (lines[0]).split(",")

worksheet.write(0, 0, words[1])
worksheet.write(0, 1, words[2])
worksheet.write(0, 2, words[3])
worksheet.write(0, 3, words[4])
worksheet.write(0, 4, words[5])
worksheet.write(0, 5, "Result")
worksheet.write(0, 6, "Execution time")
  
i = 0
while i < len(lines)-1:
    col = 0
    a1 = arr1[i]
    worksheet.write(i*10 +2, col, a1)
    col += 1
    
    a2 = arr2[i]
    worksheet.write(i*10 +2, col, a2)
    col += 1
    
    a3 = arr3[i]
    #if a3=="RoCloseDate":
     #   tempp =1
    #else:
     #   a3 = float(a3)
      #  asdf = datetime.datetime(*xlrd.xldate_as_tuple(a3, wb.datemode))
       # print(asdf)
       
       
    
    
    if(a3 == 43169.0):
        a3 = '3/10/2018'
    
    if(a3 == 43168.0):
        a3 = '3/9/2018'
        
        
    
    if(a3 == 43167.0):
        a3 = '3/8/2018'
        
    if(a3 == 43166.0):
        a3 = '3/7/2018'
    
    if(a3 == 43165.0):
        a3 = '3/6/2018'
    if(a3 == 43164.0):
        a3 = '3/5/2018'
    
    if(a3 == 43163.0):
        a3 = '3/4/2018'
    
    if(a3 == 43162.0):
        a3 = '3/3/2018'
        
    if(a3 == 43161.0):
        a3 = '3/2/2018'
        
    if(a3 == 43160.0):
        a3 = '3/1/2018'
    
    worksheet.write(i*10 +2, col, a3)
    col += 1
    
    
    a4 = arr4[i]
    worksheet.write(i*10 +2, col, a4)
    col += 1
    
    a5 = arr5[i]
    worksheet.write(i*10 + 2, col, a5)
    col += 1
    
    #quer = getQuery(a1,a2,a3,a4,a5)
    
    #days_file = open("query.txt",'r')
    
    #queee = days_file.read()
    
    start = time.time()
    
        
    end = time.time()
    
    execTime = exec_time[i]
    
    res = result[i] 
        
    
    worksheet.write(i*10 +2, col, res)
    col += 1
    
    
    worksheet.write(i*10 +2, col, execTime)
    col += 1
    
    
    
    i+=1
    
    
    
workbook.close()


    
    

    
    
    
    
    
    
    