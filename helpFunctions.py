
from datetime import datetime

def create_month_dict():
    # stores sales by month using dictionary
    data = {
        "January": 0,
        "February": 0,
        "March": 0,
        "April": 0,
        "May": 0,
        "June": 0,
        "July": 0,
        "August": 0,
        "September": 0,
        "October": 0,
        "November": 0,
        "December": 0
    }
    return data

def round_dict_values(data):
    # loop dictionary to iterate for values
    for key in data:
        # rounding values using round()
        data[key] = round(data[key])

    return data

def string_to_date(string):
    # convert date from string to date object
    date_format = '%m/%d/%Y'
    date_obj = datetime.strptime(string, date_format)
    return date_obj


def add_sheet_to_wb(wb, sheet_data, sheetName):
    #create new sheet
    wb.create_sheet(sheetName)
    #new sheet ref
    new_sheet = wb[sheetName]

    for i in sheet_data:
        new_sheet.append(i)

    return  new_sheet