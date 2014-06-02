#!/usr/bin/env python
#-*- coding: utf-8 -*-
"""
Created on May 31, 2014

Description: used for well information and observations reconciling.

@author: Jianli Gu
"""

import xlrd
import xlwt

#files
originalFile = "D:/Workdata/NobleWellsWQDatabase/Noble Wells WQ Database.xlsx"
targetFile = "D:/Workdata/NobleWellsWQDatabase/ReconciledDatabase.xls"
midTxt = "D:/Workdata/NobleWellsWQDatabase/alldata.txt"

# Step1: 
# checking the the spreadsheets' column names, the column index
rBook = xlrd.open_workbook(originalFile)

nameIndex1 = []
rSheet = rBook.sheet_by_name("Core")
for col in range(rSheet.ncols):
    colName = rSheet.cell(0,col).value
    nameIndex1.append((colName, col))

sheetNames = []
for rSheet in rBook.sheets():
    sheetName = rSheet.name
    if sheetName == "Intro" or sheetName == "Well information":
        continue
    else:
        sheetNames.append(sheetName)
        print sheetName
    
    nameIndex2 = []
    for col in range(rSheet.ncols):
        colName = rSheet.cell(0,col).value
        nameIndex2.append((colName, col))
    
    for elem1 in nameIndex1:
        if elem1 not in nameIndex2:
            print elem1
    
    print "---------------------"
    
# Step2:
# Correct the problems found in step1,including:
# the column number, name, index
# ensure all spreadsheets have the same structure.
# merge all spreadsheets in one.
sheetNames = ["Core","CUMMINS EXTE","East Pony","GREELEY CRES","Mustang",
              "West Pnoy","Wells Ranch","Flowback","Transition","Produced",
              "Horizontal","Vertical"]

wBook = xlwt.Workbook(encoding="utf-8")
wSheet = wBook.add_sheet("AllData")
rowsCount = 0
for sheetName in sheetNames:
    rSheet = rBook.sheet_by_name(sheetName)
    if sheetName == "Core":
        for row in range(rSheet.nrows):
            for col in range(rSheet.ncols):
                wSheet.write(row,col,rSheet.cell(row,col).value)
        rowsCount += rSheet.nrows
    else:
        for row in range(1, rSheet.nrows):
            for col in range(rSheet.ncols):
                wSheet.write(rowsCount+row, col, rSheet.cell(row,col).value)
        rowsCount += (rSheet.nrows - 1)
wBook.save(targetFile)
    
# Step3:
# Check the results and delete the blank rows manually
# then join the Well information into the observations.
# use TXT file to test and then import the result to xls file.
rBook = xlrd.open_workbook(originalFile)
rSheet = rBook.sheet_by_name("Well information")
  
wellinfoList = []
for row in range(1,rSheet.nrows):
    rowList = []
    for col in range(rSheet.ncols):
        rowList.append(rSheet.cell(row,col).value)
    wellinfoList.append(rowList)
  
#
rBook2 = xlrd.open_workbook(targetFile)
rSheet2 = rBook2.sheet_by_name("AllData")
  
alldataList = []
for row in range(1, rSheet2.nrows):
        key = str(rSheet2.cell(row,1).value).strip()
        newRowList = [key]
        for rowList in wellinfoList:
            location = str(rowList[0]).strip()
            if key.lower() == location.lower():
                print "Key: ", key
                newRowList.extend(rowList[1:])
                print newRowList
        alldataList.append(newRowList)

#
f1 = open(midTxt, "a")
try:
    for newRowList in alldataList:
        line = ""
        for e in newRowList:
            line += str(e) + ";"
        line = line.strip(";") + "\n"
        f1.write(line)
finally:
    f1.close()
