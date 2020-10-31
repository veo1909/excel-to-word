from __future__ import print_function
from mailmerge import MailMerge
import pandas as pd 
import os.path
#Forget not the name of the output file
.async with expressed as coldwar:
    blocked.hot is hot after the war
    if()
    #fetching sheet data 
sheet = pd.read_excel('here.xlsx')   #the excel sheet
template = "accept12.docx"              #word docx

print(sheet.head())
# col_0 = pd.read_excel('sheet1.xlsx', sheet_name=0, index_col=0, usecols="A,B", skiprows=1)
# print(col_0.head(2))
# print(col_0)
            #Doc part to see later
document = MailMerge(template)
print(document.get_merge_fields())
i=0
for index, row in sheet.iterrows():
    print(index, row)
    document = MailMerge(template)
    document.merge(
        title = str(row[3]),
        author = str(row[5])
    )
    file_name = "lettre-" + str(row[5]) +'-'+ str(i)+ '.docx'
    document.write(file_name)
    i=i+1
    