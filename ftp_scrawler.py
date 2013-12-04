import csv
import os
#import pandas as pd
from ftplib import FTP
from bs4 import BeautifulSoup as bs
from datetime import datetime
import win32com.client as win32

# Now download the files we need from FTP server, we need three files Commission_Details and Daily_Statement

def download():
    abs_path=os.getcwd()
    ftp=FTP('ftp3.interactivebrokers.com')
    ftp.login('***','****')
    ftp.cwd('****')
    buf_size=1024
    Commission_Details=[]
    Daily_Statement=[]
    Cash_Report=[]
    for fil in ftp.nlst():
        if fil.find('Commission_Details') !=-1:
            Commission_Details.append(fil)
        if fil.find('Daily_Statement') !=-1:
            Daily_Statement.append(fil)
        if fil.find('Cash_Report') !=-1:
            Cash_Report.append(fil)

    for filename in Commission_Details:
        file_handler=open(abs_path+'\\'+'Download\Commission_Details\\'+filename, 'wb')
        ftp.retrbinary("RETR "+filename,file_handler.write,buf_size)
        file_handler.close()
    print 'Finished downloading Commission_Details'

    for filename in Daily_Statement:
        file_handler=open(abs_path+'\\'+'Download\Daily_Statement\\'+filename, 'wb')
        ftp.retrbinary("RETR "+filename,file_handler.write,buf_size)
        file_handler.close()
    print 'Finished downloading Daily_Statements'

    for filename in Cash_Report:
        file_handler=open(abs_path+'\\'+'Download\Cash_Report\\'+filename, 'wb')
        ftp.retrbinary("RETR "+filename,file_handler.write,buf_size)
        file_handler.close()

Update=raw_input('Do you want to update these log files? (Y/N)')
if Update=='Y' or Update=='y':
    download()
else:
    print 'Start without updating log files.'

date430=[]
totals430=[]
date406=[]
totals406=[]
abs_path=os.getcwd()+'\Download\\'
filelist=os.listdir(abs_path+'Commission_Details')

print 'Begin reading Commission files'
for files in filelist:
    total=0
    AccountID=files[:8]
    sh=open(abs_path+'Commission_Details'+'\\'+files)
    sh.readline()
    reader=csv.reader(sh)
    for row in reader:
        line=map(float, row[11:17])
        total+=sum(line)
    
    if AccountID=='U*****':
        totals430.append(total)
        date430.append(files[-12:-4])
    else:
        totals406.append(total)
        date406.append(files[-12:-4])

print 'Begin reading NAV.'

filelist=os.listdir(abs_path+'Daily_Statement')

NAV_U****=[]
NAV_U****=[]

for files in filelist:
    
    AccountID=files[:8]
    f1=open(abs_path+'Daily_Statement\\'+files,'r')
    soup=bs(f1.read())
    s=soup('div', {'id':'tblEquitySummary_%sDivContainer' % AccountID})
    x=s[0]('tr')[2]('td',{'align':'right'})[0].contents
    
    x=x[0]
    f1.close()
    
    x=x.replace(',','')
    
    NAV=float(x)
    if AccountID=='U*****':
        NAV_U******.append(NAV)
    if AccountID=='U*****':
        NAV_U******.append(NAV)

print 'finished reading html files'

print 'begin reading cash report'

filelist=os.listdir(abs_path+'Cash_Report')
trans430=[]
trans406=[]

for files in filelist:
    total=0
    AccountID=files[:8]
    sh=open(abs_path+'Cash_Report'+'\\'+files)
    sh.readline()
    reader=csv.reader(sh)
    for row in reader:
        line=row[11]
        #print 'line is: '+line
    trans=float(line)
    if AccountID=='U******':
        trans430.append(trans)
        #date430.append(files[-12:-4])
    else:
        trans406.append(trans)
        #date406.append(files[-12:-4])

print 'Finished reading cash report.'

Pnl_U****=[]
PnlPercnt_U****=[]
Pnl_U*****.append('NA')
PnlPercnt_U*****.append('NA')
i=0
IncptToDateU*****=['NA']
IncptToDate=0
MTD_return=0
MTD_U******=['NA']
while i<len(NAV_U*****)-1:
    i+=1
    
    if datetime.strptime(date430[i],'%Y%m%d').month-datetime.strptime(date430[i-1],'%Y%m%d').month!=0:
        MTD_return=0
    Pnl_U*****.append(NAV_U*****[i]-NAV_U*****[i-1]-trans430[i])
    PnlPercnt_U*******.append(Pnl_U******[i]/(NAV_U*****[i-1]-trans430[i-1]))
    MTD_return+=Pnl_U*****[i]
    MTD_U*****.append(MTD_return)
    IncptToDate+=PnlPercnt_U*****[i]
    IncptToDateU*****.append(IncptToDate)


Pnl_U******=[]
PnlPercnt_U******=[]
Pnl_U******=['NA']
PnlPercnt_U*****=['NA']
i=0
IncptToDateU*****=['NA']
IncptToDate=0
MTD_return=0
MTD_U******=['NA']
while i<len(NAV_U******)-1:
    i+=1
    if datetime.strptime(date406[i],'%Y%m%d').month-datetime.strptime(date406[i-1],'%Y%m%d').month !=0:
        MTD_return=0
    Pnl_U******.append(NAV_U*****[i]-NAV_U******[i-1]-trans406[i])
    PnlPercnt_U*****.append(Pnl_U*****[i]/(NAV_U***[i-1]-trans406[i-1]))
    MTD_return+=Pnl_U*****[i]
    MTD_U*****.append(MTD_return)
    IncptToDate+=PnlPercnt_U******[i]
    IncptToDateU********.append(IncptToDate)

print 'Now process the format.'
for i in range(len(NAV_U****)):
    if i==0:continue
    NAV_U******[i]='{:,.2f}'.format(NAV_U******[i])
    Pnl_U******[i]='{:,.2f}'.format(Pnl_U******[i])
    MTD_U******[i]='{:,.2f}'.format(MTD_U******[i])
    PnlPercnt_U******[i]='{0:.2f}%'.format(PnlPercnt_U******[i]*100)
    IncptToDateU******[i]='{0:.2f}%'.format(IncptToDateU******[i]*100)
    totals430[i]='{:,.2f}'.format(float(totals430[i]))


for i in range(len(NAV_U******)):
    if i==0:continue
    NAV_U*****[i]='{:,.2f}'.format(NAV_U*****[i])
    Pnl_U******[i]='{:,.2f}'.format(Pnl_U******[i])
    MTD_U******[i]='{:,.2f}'.format(MTD_U******i])
    PnlPercnt_U******[i]='{0:.2f}%'.format(PnlPercnt_U*****[i]*100)
    IncptToDateU*****[i]='{0:.2f}%'.format(IncptToDateU******[i]*100)
    totals406[i]='{:,.2f}'.format(float(totals406[i]))



print 'Now I want to use win32com to write date, NAV, Pnl, PnlPercent, MTD_return and InceptToDate into target Excel file.'
excel=win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
workbook=excel.Workbooks.Open(os.path.join(os.getcwd(),'MG Daily Performance Metrics.xlsx'))
U******sheet=workbook.Worksheets('U******account')
U*****sheet=workbook.Worksheets('U***** account')

for row in range(2,len(date430)+2):
    U******sheet.Cells(row, 1).Value=date430[row-2]
    #print row, len(NAV_U******)
    U******sheet.Cells(row, 2).Value=NAV_U******[row-2]
    U******sheet.Cells(row, 3).Value=Pnl_U******[row-2]
    U******sheet.Cells(row, 4).Value=PnlPercnt_U******[row-2]
    U******sheet.Cells(row, 5).Value=MTD_U******[row-2]
    U******sheet.Cells(row, 6).Value=IncptToDateU******[row-2]
    U******sheet.Cells(row, 7).Value=trans430[row-2]
    U******sheet.Cells(row,10).Value=totals430[row-2]

for row in range(2,len(date406)+2):
    U******sheet.Cells(row, 1).Value=date406[row-2]
    U*****sheet.Cells(row, 2).Value=NAV_U***[row-2]
    U******sheet.Cells(row, 3).Value=Pnl_U****[row-2]
    U*****sheet.Cells(row, 4).Value=PnlPercnt_U*****[row-2]
    U******sheet.Cells(row, 5).Value=MTD_U*****[row-2]
    U*****sheet.Cells(row, 6).Value=IncptToDateU*****[row-2]
    U******sheet.Cells(row, 7).Value=trans406[row-2]
    U*****sheet.Cells(row,10).Value=totals406[row-2]

print 'Finished writing, now saving workbook'
workbook.Save()
excel.Application.Quit()
print 'Done!'







#dfU******=pd.DataFrame({'U****** NAV':pd.Series(NAV_U******, index=date430),'Commission Total':pd.Series(totals430, index=date430)})

#dfU*****=pd.DataFrame({'U*****NAV':pd.Series(NAV_U*****, index=date406),'Commission Total':pd.Series(totals406, index=date406)})




#dfU******.to_csv('Result\U****** Commission Total.csv')
#dfU******.to_csv('Result\U****** Commission Total.csv')





