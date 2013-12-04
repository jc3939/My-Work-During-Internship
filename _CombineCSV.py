# -*- coding: utf-8 -*-
"""
Data RankList Integration

Created on Mon Oct 05 12:19:36 2013

@author: Jianjun Chen
"""

import os
import pandas as pd
import cStringIO
import xlrd
import csv
import copy

long_shorts = []
long_short_results = []

file_list = os.listdir('.')
for each_file in file_list:
    if each_file.endswith('Long.csv') or each_file.endswith('Short.csv'):
        long_shorts.append(each_file)
    elif each_file.endswith('Long Results.csv') or each_file.endswith('Short Results.csv'):
        long_short_results.append(each_file)

Date=[]
No=[]
Symbol=[]
Name=[]
Score=[]
Amp_Id=[]
Long_Shorts=[]

for each_csv in long_shorts:
    files=open(each_csv)
    
    for row in files:
        #count+=1
        line=row.split(',')
        Date.append(line[0])
        No.append(line[1])
        Name.append(line[3])
        Score.append(line[4])
        Symbol.append(line[2])
        Amp_Id.append(each_csv[0:3])
        Long_Shorts.append(each_csv[4:9])

with open("_STOCKS_.csv",'wb') as outfile:
    writer=csv.writer(outfile)
    header=True
    for row in range(len(Date)):
        if header==True:
            writer.writerow(['Date','No','Symbol','Name','Score','Amp_Id','Long_Short'])
            header=False
        writer.writerow([Date[row],No[row],Symbol[row],Name[row],Score[row],Amp_Id[row],Long_Shorts[row]])


Date=[]
F1=[]
F2=[]
F3=[]
F4=[]
F5=[]
F6=[]
F7=[]
F8=[]
F9=[]
F10=[]
Amp_Id=[]
Long_Short=[]
for each_csv in long_short_results:
    files=open(each_csv)
    files.readline()
    for row in files:
        line=row.split(',')
        Date.append(line[0])
        F1.append(line[1])
        F2.append(line[2])
        F3.append(line[3])
        F4.append(line[4])
        F5.append(line[5])
        F6.append(line[6])
        F7.append(line[7])
        F8.append(line[8])
        F9.append(line[9])
        F10.append(line[10])
        Amp_Id.append(each_csv[0:3])
        Long_Short.append(each_csv[4:9])

with open("_RESULTS_.csv",'wb') as outfile:
    writer=csv.writer(outfile)
    header=True
    for row in range(len(Date)):
        if header==True:
            writer.writerow(['Date','F1','F2','F3','F4','F5','F6','F7','F8','F9','F10','Amp_Id','Long_Short'])
            header=False
        writer.writerow([Date[row],F1[row],F2[row],F3[row],F4[row],F5[row],F6[row],F7[row],F8[row],F9[row],F10[row].replace('\n',''),Amp_Id[row],Long_Short[row]])
        

print("Combination Done!")
#quit()
with open('_RESULTS_.csv','rb') as csvfile:
    csvfile.readline()
    MasterList=[]
    fractile=0
    while fractile<10:
        fractile+=1
        #print fractile
        csvfile.seek(0)
        for row in csvfile:
            #csvfile.seek(0)
            line=row.split(',')
            line[11]=line[11].replace('\r\n','')
            MasterList.append([line[0],line[fractile],fractile, line[11],line[12],1])

MasterList.sort(key=lambda l:(l[0],l[2]))
#print MasterList[0]
with open('_RESULTS_.csv','wb') as outfile:
    writer=csv.writer(outfile)
    header=True
    for row in MasterList:
        if header==True:
            writer.writerow(['Date','Return','Fractile','Amp_Id','Long_Short','zFractile_Id'])
            header=False
        row[4]=row[4].replace('\r\n','')
        writer.writerow(row)

