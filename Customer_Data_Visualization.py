import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.chart import (
    LineChart,
    Reference,
)
import os
from openpyxl.chart.axis import DateAxis
import pandas as pd
import numpy as np
from collections import Counter
from datetime import date

# DONE visualize data into date line chart https://openpyxl.readthedocs.io/en/stable/charts/line.html
# DONE line 62 linechart_fw_creation_date.y_axis.crossAx ????
# DONE split error in line 26
# DONE make efficient with ws
# DONE rename file
# TODO try out df.aggregate
# TODO def visulaize(col_name, is_month, chart_type)
# TODO edit functions to include 'if is_month'


# STEP BY STEP
# DONE get raw data 
# DONE aggregate data into usable format
#   DONE line chart by day)
#       format into two cols, date and how many fw created on each date
#   DONE line chart by month)
#       format date data into mm/yyyy - testing on testpage.py
#       format data into list within list containing [date, count]
#       append into ws
#       chart
# visualization:
#   DONE line chart for firewall creation date by day
#   bar chart ^^^ by month
#   multilevel donut chart for mode and nettype within mode

### note: put visualizations and data onto same sheet

# basic_user_registration_data  self disclosed data
# user_viewing_behavior_data  uid, time, like, share, comment, topic
# user_follow_behavior_data   uid, following, follow_by
# heuristic_user_data         

filepath = r'C:\Users\miche\OneDrive\Documents\CodingwDad\dad_intern\py-xl-chart.xlsx'
dict_df = pd.read_excel(io=filepath, sheet_name=['Data'], usecols=[1,36,37]) # import xlsx file, Data sheet, columns for mode, fwcreatedate, and nettype
df = dict_df['Data']

df['FwCreateMonth'] = df['FwCreateDate'].str.slice(0,7)

# Creates new workbook and retitles ws1 to 'Raw', creates 2 other sheets
data_v = Workbook()
ws1 = data_v.active
ws1.title = 'Raw'
ws2 = data_v.create_sheet('Formatted')
ws3 = data_v.create_sheet('Visualization')

# reset worksheet variables to reference sheets
ws1 = data_v['Raw']
ws2 = data_v['Formatted']
ws3 = data_v['Visualization']

# 3 seperate columns for mode, fwcreatedate, and nettype
mode = df.loc[:, 'Mode']
# fw_create_date = df.loc[:, 'FwCreateDate']
fw_create_date = df.loc[:, 'FwCreateMonth']
net_type = df.loc[:,'NetType']

# creates sorted_count_date list within list [date, count], adds header in formatted_date_data
count_date = Counter(list(fw_create_date))
count_date.pop(np.nan, None) # pop np.nan from data, if nonexistent then do nothing

# dictionary comp ONLY if the data is in YYYY-mm, NOT YYYY-mm-dd. Assigning false day value to YYYY-mm to use date function later
if list(count_date.keys())[0].count('-') == 1:
    count_date = {
     (k + '-01') : v # what we are doing
     for k, v in count_date.items() # for which items
     if k is not np.nan # conditional # use is not for np.nan because we don't know if unknowns == unknowns, but we DO know if something IS NOT unknown
     }
list_count_date = count_date.items()

sorted_count_date = sorted([[date(*[int(item) for item in k.split('-')]),v] #what we are doing to the items
                            for k,v in count_date.items() #where we find the items from
                            if k is not np.nan] #conditional filter
                            )
count_date_headers = [['Date','Count']]
formatted_date_data = count_date_headers + sorted_count_date

# appends formatted_date_data to ws2 only if ws2 is empty (to stop it from appending again every time the code is run)
if ws2['A1'].value == None: # check ws2 is empty
    for row in formatted_date_data:
        ws2.append(row)

# Line Chart for Frequency of Firewall Creation by Date
linechart_fw_creation_date = LineChart()
linechart_fw_creation_date.title = 'Frequency of Firewall Creation by Date'
linechart_fw_creation_date.style = 12
linechart_fw_creation_date.y_axis.title = 'Firewalls Created'
linechart_fw_creation_date.y_axis.crossAx = 500 # wth does this even mean
linechart_fw_creation_date.x_axis = DateAxis(crossAx=100)
linechart_fw_creation_date.x_axis.number_format = 'd-mmm'
linechart_fw_creation_date.x_axis.majorTimeUnit = 'days'
linechart_fw_creation_date.x_axis.title = 'Date'

# where to reference data from
data = Reference(ws2,
                min_col=2,
                min_row=2,
                max_col=2,
                max_row=ws2.max_row,
                )
dates = Reference(ws2,
                  min_col=1,
                  min_row=2,
                  max_col=1,
                  max_row=ws2.max_row,
                  )
linechart_fw_creation_date.add_data(data, titles_from_data=True)
linechart_fw_creation_date.set_categories(dates)

ws2.add_chart(linechart_fw_creation_date, 'C1')
data_v.save(os.path.join('Data_Intern', 'dv_intern.xlsx'))



# counter_ex = Counter(['2020-12-12','2020-12-12','2034-23-03'])
# print(counter_ex)

# df.to_excel(excel_writer='temp1.xlsx')



# print(df)












