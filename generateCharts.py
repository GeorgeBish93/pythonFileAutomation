import calendar as cl
from openpyxl.chart import BarChart,PieChart,LineChart,Reference
def sales_line_chart(wb,new_sheet):
    #create the chart and the references
    chart = LineChart()
    data = Reference(new_sheet, min_col=2, max_col=4, min_row=2, max_row=2)
    labels = Reference(new_sheet, range_string="chart!A1:M1")

    #set the title, add the data to the chart and set the categories
    chart.title = 'Customer Payment Method Distribution'
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)

    ##cosmetics
    chart.title = 'Sales By Month 2019'
    chart.x_axis.title = 'Month'
    chart.y_axis.title = 'Total Sales (USD Mil)'

    #add the chart and save the workbook as a new file
    new_sheet.add_chart(chart, 'F5')
