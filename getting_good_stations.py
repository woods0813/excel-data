import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt


data_dir = r'C:\Users\Tommy\Downloads\\'

#reading in excel file
name_filename = 'needed_cities.xlsx'
data_filename = 'My Pillow Quarterly-Monthly Recap 3-13-23 Ann.xlsx'

data_file = pd.read_excel(os.path.join(data_dir, data_filename), 'Quarterly Recap')
name_file = pd.read_excel(os.path.join(data_dir, name_filename), header = None)

#gathering necessary product names
city_names = []
product_names = ['GIZA SHEETS', 'MyPillow', 'Slippers', 'Bedding Multiproduct', 'GizaSheets', 'MyPillowMIX', 'Percale Sheets']

for i in range(len(name_file[0])):
    city_names.append(name_file[0][i])
    
#parsing through the excel data to take only the desired cities
data_labels = []
labels_key = data_file.keys()[0]
numbers_key = data_file.keys()[1]
city_idx = [0]*len(city_names)
city_cnt = [0]*len(city_names)

for i in range(len(data_file[labels_key])):
    string_label = data_file[labels_key][i]
    if str(type(string_label)).find('str') == -1:
        string_label = str(string_label)
    for n in range(len(city_names)):
        name = city_names[n]
        
        if name.find(string_label) != -1 or string_label.find(name) != -1 or string_label.lower().find(name.lower()) != -1 or name.lower().find(string_label.lower()) != -1:
            city_idx[n] = i
            city_cnt[n] += 1
            
sorted_city_idx = sorted(city_idx)
orig_city_idx = []
sorted_city_names = []

for idx in sorted_city_idx:
    orig_city_idx.append(city_idx.index(idx))
    sorted_city_names.append(city_names[city_idx.index(idx)])


#write the corresponding output to excel
good_stations_filename = 'Good_Products_Big_Cities.xlsx'
good_stations_file = os.path.join(data_dir, good_stations_filename)
writer = pd.ExcelWriter(good_stations_file,engine='xlsxwriter')   
workbook=writer.book
worksheet=workbook.add_worksheet('Validation')
writer.sheets['Validation'] = worksheet
 
current_idx = 0
for df in df_list:
    df.to_excel(writer,sheet_name='Validation',startrow=1 , startcol=1 + current_idx, index = False)  
    current_idx += 1

writer.close()


##NATIONAL CABLE##

data_dir = r'C:\Users\Tommy\Downloads\\'

data_filename = 'My Pillow Quarterly-Monthly Recap 3-13-23 CABLE Ann.xlsx'

data_file = pd.read_excel(os.path.join(data_dir, data_filename), 'Quarterly Recap')
city_names = ['National Cable']
product_names = ['GIZA SHEETS', 'MyPillow', 'Slippers', 'Bedding Multiproduct', 'GizaSheets', 'MyPillowMIX', 'Percale Sheets']

data_labels = []
labels_key = data_file.keys()[0]
numbers_key = data_file.keys()[1]
city_idx = [0]*len(city_names)
city_cnt = [0]*len(city_names)

for i in range(len(data_file[labels_key])):
    string_label = data_file[labels_key][i]
    if str(type(string_label)).find('str') == -1:
        string_label = str(string_label)
    for n in range(len(city_names)):
        name = city_names[n]
        
        if name.find(string_label) != -1 or string_label.find(name) != -1 or string_label.lower().find(name.lower()) != -1 or name.lower().find(string_label.lower()) != -1:
            city_idx[n] = i
            city_cnt[n] += 1
            

df_list = []

for i in range(len(city_names)):
    new_df = pd.DataFrame(columns=[sorted_city_names[i]])
    if i < len(city_names) - 1:
        station_segment = data_file[labels_key][sorted_city_idx[i] + 1: sorted_city_idx[i + 1]].tolist()
        numbers_segment = data_file[numbers_key][sorted_city_idx[i] + 1: sorted_city_idx[i + 1]].tolist()
    else:
        station_segment = data_file[labels_key][sorted_city_idx[i] + 1:].tolist()
        numbers_segment = data_file[numbers_key][sorted_city_idx[i] + 1:].tolist()
   
    good_stations_list = []
    good_stations_cnt = 0
    for n in range(len(station_segment)):
        if station_segment[n] not in product_names:
            if len(good_stations_list) > 0 and good_stations_cnt == 0:
                good_stations_list.remove(station_name)
            good_stations_cnt = 0
            station_name = station_segment[n]
            good_stations_list.append(station_name)
        else:
            if numbers_segment[n] >= 1:
                good_stations_cnt += 1
                good_stations_list[-1] += ', ' + station_segment[n]
        
    if len(good_stations_list) > 0 and good_stations_cnt == 0:
        good_stations_list.remove(station_name)
        
    new_df[sorted_city_names[i]] = good_stations_list
    
    df_list.append(new_df)

good_stations_filename = 'Good_Products_National_Cable.xlsx'
good_stations_file = os.path.join(data_dir, good_stations_filename)
writer = pd.ExcelWriter(good_stations_file,engine='xlsxwriter')   
workbook=writer.book
worksheet=workbook.add_worksheet('Validation')
writer.sheets['Validation'] = worksheet
 
current_idx = 0
for df in df_list:
    df.to_excel(writer,sheet_name='Validation',startrow=1 , startcol=1 + current_idx, index = False)  
    current_idx += 1

writer.close()