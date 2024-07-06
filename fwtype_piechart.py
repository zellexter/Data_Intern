import os
from Customer_Data_Visualization import DataVisualization

filepath = r'C:\Users\miche\OneDrive\Documents\CodingwDad\Data_Intern\py-xl-chart.xlsx'
save_file = os.path.join('Data_Intern', 'net_type.xlsx')
dv = DataVisualization(input_file=filepath, output_file=save_file)

# TODO create net type data visualization piechart
# link: https://openpyxl.readthedocs.io/en/stable/charts/pie.html
#   TODO create general pie chart for mode types
#   TODO create projected pie chart for each mode type, for each nettype

dv.main()
