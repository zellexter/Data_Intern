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
# DONE def visulaize(col_name, is_month, chart_type)
# DONE edit functions to include 'if is_month'
# TODO save into different worksheets
#   line 50 add new arg into for loop
#   format_data() wb to ws


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

def main():
    filepath = r'C:\Users\miche\OneDrive\Documents\CodingwDad\dad_intern\py-xl-chart.xlsx'
    df = import_data(filepath)
    for col_name, output_file in [
        ('FwCreateDate', 'dv_intern_daily.xlsx'),
        ('FwCreateMonth', 'dv_intern_month.xlsx'),
    ]:
        wb = prepare_wb()
        format_data(df, wb, col_name)
        linechart_fw_creation_date = draw_line_chart(wb)
        save_chart(wb, linechart_fw_creation_date, output_file)
    


def import_data(filepath):
    dict_df = pd.read_excel(io=filepath, sheet_name=['Data'], usecols=[1,36,37]) # import xlsx file, Data sheet, columns for mode, fwcreatedate, and nettype
    df = dict_df['Data']

    df['FwCreateMonth'] = df['FwCreateDate'].str.slice(0,7)

    return df


def prepare_wb():
    # Creates new workbook and retitles ws1 to 'Raw', creates 2 other sheets
    wb = Workbook()
    ws = wb.active
    ws.title = 'Formatted'
    return wb


def format_data(df, wb, col_name):
    '''This function creates a visualization for date data'''
    
    # 3 seperate columns for mode, fwcreatedate, and nettype
    mode = df.loc[:, 'Mode']
    fw_create_date = df.loc[:, col_name]
    net_type = df.loc[:,'NetType']

    # creates sorted_count_date list within list [date, count], adds header in formatted_date_data
    count_date = Counter(list(fw_create_date))
    count_date.pop(np.nan, None) # pop np.nan from data, if nonexistent then do nothing
    # definition of when is_month is True
    is_month = list(count_date.keys())[0].count('-') == 1
    # what to do when is_month is True
    if is_month:
        count_date = {
            (k + '-01') : v # what we are doing
            for k, v in count_date.items() # for which items
            if k is not np.nan # conditional # use is not for np.nan because we don't know if unknowns == unknowns, but we DO know if something IS NOT unknown
            }
    # list_count_date = count_date.items()
        
    sorted_count_date = sorted([[date(*[int(item) for item in k.split('-')]),v] #what we are doing to the items
                            for k,v in count_date.items() #where we find the items from
                            if k is not np.nan] #conditional filter
                            )
    count_date_headers = [['Date','Count']]
    formatted_date_data = count_date_headers + sorted_count_date
    # return formatted_date_data

    ws2 = wb['Formatted']
    # appends formatted_date_data to ws2 only if ws2 is empty (to stop it from appending again every time the code is run)
    if ws2['A1'].value == None: # check ws2 is empty
        for row in formatted_date_data:
            ws2.append(row)
        

def draw_line_chart(wb):
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
    ws2 = wb['Formatted']
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

    return linechart_fw_creation_date


def save_chart(wb, linechart_fw_creation_date, output_file):
    ws2 = wb['Formatted']
    ws2.add_chart(linechart_fw_creation_date, 'C1')
    wb.save(os.path.join('Data_Intern', output_file))


if __name__ == '__main__':
    main()











