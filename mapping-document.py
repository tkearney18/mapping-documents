import os
import xlrd
from xlutils.copy import copy
import csv

path = 'sample docs/'            
all_files = os.listdir(path) 

#print(all_files)

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
        
docMap = []
i = 0
for materialId in materialIdList:
    index = str(materialId).rstrip('0').rstrip('.')
    docItem = {'Material ID': index, 'Screen Name': fileNameList[i], 'Contract ID': contractIdList[i]}
    i += 1
    docMap.append(docItem)
print(docMap)
for listItem in docMap:
    indices = [i for i, f in enumerate(all_files) if listItem.get('Material ID') in f]
    if indices:
        listItem['File Name'] = all_files[indices[0]]
    else:
        listItem['File Name'] = ''
    #if listItem.get('Material ID') in all_files:
        #print(listItem.get('Material ID'))
    # for doc in all_files:
    #     print(listItem.get('Material ID'))
    #     if listItem.get('Material ID') in doc:
    #         listItem['File Name'] = doc
    #     else:
    #         listItem['File Name'] = ''
#print(docMap)
keys = docMap[0].keys()
with open('anyting.csv', 'w') as outputFile:
    #dictWriter = csv.writer(outputFile)
    dictWriter = csv.DictWriter(outputFile,keys,lineterminator='\n')
    dictWriter.writeheader()
    dictWriter.writerows(docMap)