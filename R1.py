import pylightxl
#import pandas as pd
import glob
import xlsxwriter
#import numpy as np

path = glob.glob("*.xlsx")


workbook = xlsxwriter.Workbook('Combined.xlsx')
worksheet = workbook.add_worksheet()

array = []
array2=[]



for f in path:
    stb_array=[]
    
    db = pylightxl.readxl(f)
    sheet1 = db.ws_names[0]
    sheet2 = db.ws_names[1]

    # sheet2=db.ws_names[1]
    print(sheet1)
    print(sheet2)
    # print(sheet2)
 
    for row in db.ws(ws=sheet1).rows:
        row.append(sheet1)
        array.append(row)
    
    for row in db.ws(ws=sheet2).rows:
        if(row[1]!="Break" and row[1]!="Lunch" and row[1]!="Debrief") :
          row.append(sheet2)
          stb_array.append(row)
    
    array2.append(stb_array)

print(len(array2))

def arrayreverse(A, n, p):
   i = 3  
   while(i<n):
      L = i 
      R = min(i + p - 1, n - 1) 
      while (L < R):
         A[L], A[R] = A[R], A[L]
         L+= 1;
         R-+1
      i+= p


for i in array:
    if(i[1]=="Break" or i[1]=="Lunch" or i[1]=="Debrief"):
        array.pop(array.index(i))


array3=[]

for j in array2:
    arrayreverse(j,len(j),2)
   
for i in array2: 
    for k in range(0,len(i)):
      array3.append(i[k])

 

row = 0



for col, data in enumerate(array):
    worksheet.write_row(col, row, data)

for col, data in enumerate(array3):
    worksheet.write_row(col, row+12, data)


workbook.close()



# a=db.ws(ws='Sheet1').index(row="1, col=1)
