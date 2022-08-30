import pylightxl
import glob
import xlsxwriter

path = glob.glob("*.xlsx")


workbook = xlsxwriter.Workbook('CombinedR1.xlsx')
worksheet = workbook.add_worksheet()

array = []
array2=[]
array4=[]



for f in path:
    stb_array=[]
    
    db = pylightxl.readxl(f)
    sheet1 = db.ws_names[0]
    sheet2 = db.ws_names[1]

    count = 0
    
    for row in db.ws(ws=sheet1).rows:
        if(row[1] != "Break" and row[1] != "Lunch" and row[1] != "Debrief" and row[1]):
            count += 1
            if(path.index(f) == 0):
                 array.append(row)
                 if(count > 5):
                       row.append(sheet1)
            else:
                if(count > 3):
                    array.append(row)
                    row.append(sheet1)


    count=0
    for row in db.ws(ws=sheet2).rows:
            if(row[1] != "Break" and row[1] != "Lunch" and row[1] != "Debrief" and row[1]):
                count += 1
                if(path.index(f) == 0):
                    row.append("")
                    row.append(sheet2)
                    stb_array.append(row)
                else:
                    if(count > 3):
                        row.append("")
                        row.append(sheet2)
                        stb_array.append(row)
    array2.append(stb_array)


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


def arrayreverse2(A, n, p):
   i = 0 
   while(i<n):
      L = i 
      R = min(i + p - 1, n - 1) 
      while (L < R):
         A[L], A[R] = A[R], A[L]
         L+= 1;
         R-+1
      i+= p

array3=[]

for j in array2:
    if(array2.index(j)==0):
        arrayreverse(j,len(j),2)
    else:
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
