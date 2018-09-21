import os
import xlrd
from xlutils.copy import copy
import csv

path = 'sample docs/'            
all_files = os.listdir(path) 

print(all_files)

wbName = 'Plan Doc Map.xlsx'
sheetName = 'Sheet1'

rb = xlrd.open_workbook(wbName)
sh = rb.sheet_by_name(sheetName)
row = sh.row(0)
for colidx, cell in enumerate(row):
    if cell.value == "Contract ID":
        contractIdList = sh.col_values(colidx, 1)
    elif cell.value == "Material ID":
        materialIdList = sh.col_values(colidx, 1)
    elif cell.value == "File Name":
        fileNameList = sh.col_values(colidx, 1)
        
docMap = {}
for materialId in materialIdList:
    for fileName in fileNameList:
        index = str(materialId).rstrip('0').rstrip('.')
        docMap[index] = [fileName]

for listItem in docMap:
    for doc in all_files:
        if listItem in doc:
            docMap[str(materialId).rstrip('0').rstrip('.')].append(doc)  
print(docMap)
with open('anyting.csv', 'w') as outputFile:
    #dictWriter = csv.writer(outputFile)
    dictWriter = csv.DictWriter(outputFile,keys)
    dictWriter.writeheader()
    dictWriter.writerows(docList)