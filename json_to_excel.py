##directories and loading image and annotation data from jsons
import json
import xlwt
from xlwt import Workbook

base_directory=r'C:\Users\Tommy\AppData\Local\Programs\Python\Python37\Fashionpedia\\'
train_instances_directory=base_directory + 'instances_attributes_train2020.json'
categories=[]

with open(train_instances_directory,'r') as train_json_file:
    train_json=json.load(train_json_file)
    
for i in range(len(train_json['categories'])):
    categories.append(train_json['categories'][i]['name'])

wb=Workbook()

sheet1=wb.add_sheet('Sheet 1')

for i in range(len(categories)):
    sheet1.write(0, i,categories[i])
    
wb.save('json_to_xl_example.xls')
