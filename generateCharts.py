import calendar as cl
from openpyxl.chart import BarChart,PieChart,LineChart,Reference


def payment_bar_chart(wb,new_sheet):
    #create the chart and the references
    chart = BarChart()
    data = Reference(new_sheet, min_col=1, max_col=4 ,min_row=1, max_row=3)
    labels = Reference(new_sheet, min_col=1, min_row=2, max_row=3)

    #add the data to the chart and set the categories
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)

    ##cosmetics
    chart.title = 'Payment Methods by Gender'
    chart.x_axis.title = 'Payment Method'
    chart.y_axis.title = 'Method Usage Percentage'

    #add the chart and save the workbook as a new file
    new_sheet.add_chart(chart, 'F2')
def category_bar_chart(wb,new_sheet):
    #create the chart and the references
    chart = BarChart()
    data = Reference(new_sheet, min_col=2, min_row=1, max_row=7)
    labels = Reference(new_sheet, min_col=1, min_row=2, max_row=7)

    #add the data to the chart and set the categories
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)

    ##cosmetics
    chart.title = 'Sales by Categories - Bar Chart'
    chart.x_axis.title = 'Categories'
    chart.y_axis.title = 'Total Sales (USD Dollar)'
    chart.legend = None

    #add the chart and save the workbook as a new file
    new_sheet.add_chart(chart, 'F2')


def category_pie_chart(wb,new_sheet):
    #create the chart and the references
    chart = PieChart()
    data = Reference(new_sheet, min_col=2, min_row=1, max_row=7)
    labels = Reference(new_sheet, min_col=1, min_row=2, max_row=7)

    #add the data to the chart and set the categories
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)

    ##cosmetics
    chart.title = 'Sales by Categories - Pie Chart'

    #add the chart and save the workbook as a new file
    new_sheet.add_chart(chart, 'F18')

def sales_line_chart(wb,new_sheet):
    #create the chart and the references
    chart = LineChart()
    data = Reference(new_sheet, min_col=2, min_row=1, max_row=13)
    labels = Reference(new_sheet, min_col=1, min_row=2, max_row=13)

    #add the data to the chart and set the categories
    chart.title = 'Sales by Month 2019'
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(labels)

    ##cosmetics
    chart.title = 'Sales by Month (2019)'
    chart.x_axis.title = 'Month'
    chart.y_axis.title = 'Total Sales (USD Mil)'
    chart.legend = None
    #add the chart and save the workbook as a new file
    new_sheet.add_chart(chart, 'F2')

