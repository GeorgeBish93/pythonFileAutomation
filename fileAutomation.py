import openpyxl as xl
import generateCharts as gc
import helpFunctions as hf

from datetime import datetime



# function that gets file src and return a workbook object 'wb' from the loaded file
def load_workbook(filename):
    #Load
    return xl.load_workbook(filename)



def payment_method_by_gender(wb, sheet, exportFilename):
    cashF = 0
    cashM = 0
    cardF = 0
    cardM = 0
    ewalletF = 0
    ewalletM = 0

    # for lop that run on rows
    for row in range(2, sheet.max_row + 1):
        genderCell = sheet.cell(row, 5)
        paymentCell = sheet.cell(row, 13)

        if paymentCell.value == "Ewallet":
            if genderCell.value == "Female":
                ewalletF += 1
            if genderCell.value == "Male":
                ewalletM += 1
        if paymentCell.value == "Cash":
            if genderCell.value == "Female":
                cashF += 1
            if genderCell.value == "Male":
                cashM += 1
        if paymentCell.value == "Credit card":
            if genderCell.value == "Female":
                cardF += 1
            if genderCell.value == "Male":
                cardM += 1

    cashTotal = cashM+cashF
    percentageCashF = (cashF / cashTotal) * 100
    percentageCashM = (cashM / cashTotal) * 100

    cardTotal = cardF+cardM
    percentageCardF = (cardF / cardTotal) * 100
    percentageCardM = (cardM / cardTotal) * 100

    eWalletTotal = ewalletM + ewalletF
    percentageWalletF = (ewalletF / eWalletTotal) * 100
    percentageWalletM = (ewalletM / eWalletTotal) * 100

    #add new sheet
    sheet_data = (
        ("", "Cash%", "Credit Card%", "Ewallet%"),
        ("Male", percentageCashM, percentageCardM, percentageWalletM),
        ("Female", percentageCashF, percentageCardF, percentageWalletF)
    )
    new_sheet = hf.add_sheet_to_wb(wb, sheet_data, "PaymentMethodByGender")


    wb.save(exportFilename)

def sales_by_category(wb, sheet, exportFilename):
    fashion = 0
    food = 0
    electronic = 0
    sport = 0
    home = 0
    health =0

    #for lop that run on rows
    for row in range(2,sheet.max_row +1):
        categoryCell = sheet.cell(row,6)
        totalCell = sheet.cell(row,10)

        if categoryCell.value == "Fashion accessories":
            fashion += int(totalCell.value)
        if categoryCell.value == "Food and beverages":
            food += int(totalCell.value)
        if categoryCell.value == "Electronic accessories":
            electronic += int(totalCell.value)
        if categoryCell.value == "Sports and travel":
            sport += int(totalCell.value)
        if categoryCell.value == "Home and lifestyle":
            home += int(totalCell.value)
        if categoryCell.value == "Health and beauty":
            health += int(totalCell.value)


    #add new sheet
    sheet_data = (
        ("", "Electronic accessories", "Fashion accessories", "Food and beverages", "Health and beauty", "Home and lifestyle", "Sports and travel" ),
        ("Total Sales(US Dollars)", electronic, fashion, food, health, home, sport)
    )
    new_sheet = hf.add_sheet_to_wb(wb, sheet_data, "SalesByCategory")

    wb.save(exportFilename)

def sales_by_month(wb, sheet, exportFilename):
    #stores sales by month using dictionary
    data = hf.create_month_dict();

    #for lop that run on rows
    for row in range(2,sheet.max_row +1):
        totalCell = sheet.cell(row,10)
        dateCell = sheet.cell(row,11)

        #convert date from string to date object
        date_obj = hf.string_to_date(dateCell.value)

        if date_obj.month == 1:
            data["January"] = data["January"] + totalCell.value
        if date_obj.month == 2:
            data["February"] = data["February"] + totalCell.value
        if date_obj.month == 3:
            data["March"] = data["March"] + totalCell.value
        if date_obj.month == 4:
            data["April"] = data["April"] + totalCell.value
        if date_obj.month == 5:
            data["May"] = data["May"] + totalCell.value
        if date_obj.month == 6:
            data["June"] = data["June"] + totalCell.value
        if date_obj.month == 7:
            data["July"] = data["July"] + totalCell.value
        if date_obj.month == 8:
            data["August"] = data["August"] + totalCell.value
        if date_obj.month == 9:
            data["September"] = data["September"] + totalCell.value
        if date_obj.month == 10:
            data["October"] = data["October"] + totalCell.value
        if date_obj.month == 11:
            data["November"] = data["November"] + totalCell.value
        if date_obj.month == 12:
            data["December"] = data["December"] + totalCell.value

    #round data values
    data = hf.round_dict_values(data);

    #create new sheet
    sheet_data = (
        ('', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'),
        ("Total Sales(US Dollar)", data["January"], data["February"], data["March"], data["April"],
         data["May"], data["June"], data["July"], data["August"], data["September"], data["October"],
         data["November"],data["December"])
    )
    new_sheet = hf.add_sheet_to_wb(wb, sheet_data, "SalesByMonth")

    # call line chart func
    gc.sales_line_chart(wb,new_sheet)

    wb.save(exportFilename)



def net_income(wb, sheet, exportFilename):
    taxes = 0
    total = 0

    #for lop that run on rows
    for row in range(2,sheet.max_row +1):
        tax_cell = sheet.cell(row,9)
        total_cell = sheet.cell(row, 10)

        taxes += int(tax_cell.value)
        total += int(total_cell.value)

    #create new sheet
    sheet_data = (
        ("", "Total Taxes", "Total Income", "Net Income"),
        ("US Dollar",taxes,total,(total-taxes))
    )
    new_sheet = hf.add_sheet_to_wb(wb,  sheet_data, "NetIncome")

    wb.save(exportFilename)