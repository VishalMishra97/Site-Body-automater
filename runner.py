import pandas as pd
import numpy as np
from openpyxl import load_workbook
from xlwt import Workbook
import xlrd
import csv
import math
import openpyxl
import xlwt
from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy
from tempfile import TemporaryFile
def find(s,string):
    t=len(s)
    i=0
    while(i<t):
        if(str(string)!='nan'):
            if(string[0]+string[1]==s[i]):
                return i
            
        i+=1
test=pd.read_csv("path\\Codes.csv")
t=[]
k=[]
a=pd.DataFrame(test)
values=[]
for i in a:
    values.append(i)
final=[[]]
t=0
s=[]
for i in test['Code']:
        if(str(i)!='nan'):
            if(i[0]+i[1] not in s):
                s.append(i[0]+i[1])
print(s)
k=len(a)
final=[[]]
count=1
for i in values:
    if(i=="Code"):
        continue
    else:
        j=0
        t=0
        n=[]
        while(t<len(s)):
            n.append("")
            t+=1
        while(j<k):
            index=find(s,test['Code'][j])
            if(str(test[i][j])!='nan'):
                if(s[index]=='AA' or s[index]=='BB' or s[index]=='CC'):
                    n[index]+=test[i][j]+'\n'
                else:
                    n[index]+=test[i][j]
            j+=1
        rb = open_workbook("path\\random.xls")
        wb = copy(rb)
        sheet = wb.get_sheet(0)
        for i,e in enumerate(n):
            sheet.write(count,i,e)
        wb.save("path\\random.xls")
        count+=1
rb = open_workbook("path\\random.xls")
wb = copy(rb)
sheet = wb.get_sheet(0)
t=0
for i in s:
    sheet.write(0,t,i)
    t+=1
wb.save("path\\random.xls")
wb1= openpyxl.Workbook()
wb1.save('path\\output.xls')
dataread=pd.read_excel('path\\random.xls')
inputfile=pd.read_excel('path\\Input.xlsx')
col=0
for i in inputfile:
    for j in inputfile[i]:
        if(j==0):
            continue
        else:
            temper=j
            for p in s:
                temp='{'+p+'}'
                if(j.find(temp)!=-1):
                    print(j.find(temp))
                    row=2
                    for ter in dataread[p]:
                        temper=temper.replace(str(temp),str(ter))
                        print(temper)
                        tb = open_workbook("path\\tester.xls")
                        wb = copy(tb)
                        sheet = wb.get_sheet(0)
                        sheet.write(row,col,temper)
                        wb.save("path\\output.xls")
                        row+=1
                        
    col+=1
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    