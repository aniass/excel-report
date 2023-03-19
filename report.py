"""The creating automated excel report from sales data"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font 


output_path = 'C:\Python Scripts\Projects_new\Excel_project\sales_report3.xlsx'


def excel_report(file):
    '''Function to create automated excel report.
       The file name should have the following structure: your_file.xlsx'''
    
    df = pd.read_excel(file, sheet_name='Sheet1')

    # make pivot table Income by city
    income_city = df.pivot_table(index='City',
                                values='Total',
                                columns='Product line',
                                aggfunc='sum').round(0)

    # send the report table to excel file
    income_city.to_excel('sales_report3.xlsx',
                      sheet_name='Sales_city',
                      startrow=4)

    # loading workbook and selecting sheet
    wb = load_workbook('sales_report3.xlsx')
    sheet = wb.active
    sheet = wb['Sales_city']

    # change column dimension
    sheet.column_dimensions['A'].width = 12
    sheet.column_dimensions['B'].width = 20
    sheet.column_dimensions['C'].width = 20
    sheet.column_dimensions['D'].width = 20
    sheet.column_dimensions['E'].width = 20
    sheet.column_dimensions['F'].width = 20
    sheet.column_dimensions['G'].width = 20

    # Bar chart adding
    chart = BarChart()
   
    # cell references (original spreadsheet)
    min_column = wb.active.min_column
    max_column = wb.active.max_column
    min_row = wb.active.min_row
    max_row = wb.active.max_row

    data = Reference(sheet,
                 min_col=min_column+1,
                 max_col=max_column,
                 min_row=min_row,
                 max_row=max_row) 

    categories = Reference(sheet,
                       min_col=min_column,
                       max_col=min_column,
                       min_row=min_row+1,
                       max_row=max_row) 

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    sheet.add_chart(chart, "B11") 
    chart.title = 'Sales by city'
    chart.style = 42                
    chart.x_axis.title = "City"
    chart.y_axis.title = "Total sales"

    # formatting the report
    sheet['D1'] = 'Sales Report'
    sheet['D2'] = 'Sales by cities and products'
    sheet['D1'].font = Font('Arial', bold=True, size=24)
    sheet['D1'].alignment = Alignment(horizontal="center")
    sheet['D2'].font = Font('Arial', bold=True, size=10, italic = True)
    sheet['D2'].alignment = Alignment(horizontal="center")

    wb.save(output_path)
    return wb


if __name__ == '__main__':
    excel_report('supermarket_sales.xlsx')
