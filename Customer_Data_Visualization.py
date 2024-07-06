from openpyxl import Workbook, load_workbook
from openpyxl.chart import (
    LineChart,
    Reference,
    PieChart,
    ProjectedPieChart,
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
# DONE save into different worksheets
#   DONE main() add new arg into for loop that defines ws variable as prepare_ws()
#   DONE function prepare_ws() that creates new ws (like prepare_wb) with specified title (related to col_name)
#   DONE revisit prepare_wb() and edit out obsolete code
#   DONE format_data() wb to ws, appends all formatted data into specified ws instead of wb
#   DONE draw_line_chart() wb to ws, take data from specified ws instead of wb
#   DONE save_chart() wb to ws, save chart into specified ws instead of wb
# TODO generalize into class
# TODO generalize format_data() to exclude pulling specific columns
#   pass in which columns and where to put data
#   pass in type of data and how to treat


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

class DataVisualization():
    '''TODO'''

    def __init__(self, input_file, output_file) -> None:
        self.input_file = input_file
        self.output_file = output_file
        self.wb = None

    def main(self):
        '''Main Function
            Assigns dataframe into usable variable
            Prepares a workbook
            Loops through specified data columns and creates unique worksheet where visualization is stored
            Saves workbook'''
        
        df = self.import_data()
        
        self.wb = self.prepare_wb()

        # create linechart
        for col_name, output_worksheet in [
            ('FwCreateDate', 'Daily Chart'),
            ('FwCreateMonth', 'Monthly Chart'),
        ]:
            ws = self.prepare_ws(output_worksheet)
            self.format_data_linechart(df, ws, col_name)
            self.draw_linechart(ws)

        # create piechart
        for col_name, output_worksheet in [
            ('Mode', 'Mode Piechart')
        ]:
            ws = self.prepare_ws(output_worksheet)
            self.format_data_piechart(df, ws, col_name)
            self.draw_piechart(ws)

        self.save_wb()
        

    def import_data(self):
        '''This function reads an excel sheet and imports specific columns into a dataframe.
        Creates by month data column from date data. Returns the dataframe.'''

        # import xlsx file, Data sheet, columns for mode, fwcreatedate, and nettype
        dict_df = pd.read_excel(io=self.input_file, sheet_name=['Data'], usecols=[1,36,37]) 
        df = dict_df['Data']

        df['FwCreateMonth'] = df['FwCreateDate'].str.slice(0,7)

        return df


    def prepare_wb(self):
        '''This function creates a new workbook'''
        wb = Workbook()
        wb.remove(wb['Sheet'])
        return wb


    def prepare_ws(self, ws_name): # pass in col_name as name for ws
        '''This function creates a new worksheet in the workbook'''
        ws = self.wb.create_sheet(ws_name)
        return ws


    def format_data_linechart(self, df, ws, col_name):
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

        # appends formatted_date_data to ws2 only if ws2 is empty (to stop it from appending again every time the code is run)
        if ws['A1'].value == None: # check ws2 is empty
            for row in formatted_date_data:
                ws.append(row)
            
    def format_data_piechart(self, df, ws, col_name):
        '''This function creates visualization for a piechart'''

        # create column for mode data
        mode = df.loc[:, col_name]

        # creates count_date
        count_mode = Counter(list(mode))
        count_mode.pop(np.nan, None) # pop np.nan from data, if nonexistent then do nothing

        # list comp to sort mode types into usable format
        sorted_count_mode = sorted([[k,v]
                                    for k,v in count_mode.items()
                                    if k is not np.nan] # gets rid of np.nan
                                   )
        count_mode_headers = [['Mode', 'Count']]
        formatted_mode_data = count_mode_headers + sorted_count_mode
        # return formatted_mode_data

        # appends formatted_mode_data to ws2 only if ws2 is empty (to stop it from appending again every time the code is run)
        if ws['A1'].value == None: # check ws2 is empty
            for row in formatted_mode_data:
                ws.append(row)



    def format_data_projectedpiechart(self, df, ws, col_name):
        ''''''

    def draw_linechart(self, ws):
        '''This function draws a line chart and saves it into a cell'''

        # Line Chart
        linechart = LineChart()
        linechart.title = 'Frequency of Firewall Creation by Date'
        linechart.style = 12
        linechart.y_axis.title = 'Firewalls Created'
        linechart.y_axis.crossAx = 500 # wth does this even mean
        linechart.x_axis = DateAxis(crossAx=100)
        linechart.x_axis.number_format = 'd-mmm'
        linechart.x_axis.majorTimeUnit = 'days'
        linechart.x_axis.title = 'Date'
        
        # where to reference data from
        data = Reference(ws,
                        min_col=2,
                        min_row=2,
                        max_col=2,
                        max_row=ws.max_row,
                        )
        dates = Reference(ws,
                        min_col=1,
                        min_row=2,
                        max_col=1,
                        max_row=ws.max_row,
                        )
        linechart.add_data(data, titles_from_data=True)
        linechart.set_categories(dates)

        ws.add_chart(linechart, 'C1')

    def draw_piechart(self, ws):
        '''This function draws a pie chart and saves it into a cell'''
        # TODO generalize column
        # TODO show percentages
        
        piechart = PieChart()
        labels = Reference(ws, min_col=1, min_row=3, max_row=4)
        data = Reference(ws, min_col=2, min_row=2, max_row=4)
        piechart.add_data(data, titles_from_data=True)
        piechart.set_categories(labels)
        piechart.title = 'Mode Types'

        # add chart to ws
        ws.add_chart(piechart, 'C1')

    def save_wb(self):
        '''This function saves the excel file'''
        self.wb.save(self.output_file)


# if DataVisualization() is run on the original file, use specified filepath and save_file.
if __name__ == '__main__':
    filepath = r'C:\Users\miche\OneDrive\Documents\CodingwDad\Data_Intern\py-xl-chart.xlsx'
    save_file = os.path.join('Data_Intern', 'data_intern.xlsx')
    DataVisualization(input_file=filepath, output_file=save_file).main()
        # class() create an instance of the class
        # example : Eevee = Canine(gen=female, col=brown, size=m)
        #           Cobra = Canine(gen=male, col=tribrown, size=xs)











