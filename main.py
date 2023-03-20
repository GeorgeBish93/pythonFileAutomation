import fileAutomation as fa

# call loading func
filename = 'sales.xlsx'
wb = fa.load_workbook(filename)

#Results filename
exportFilename = "Results.xlsx"

# call data automation funcions
paymentsByGender = fa.payment_method_by_gender(wb, wb['Worksheet'], exportFilename)
salesByCategory = fa.sales_by_category(wb, wb['Worksheet'], exportFilename)
salesByMonth = fa.sales_by_month(wb, wb['Worksheet'], exportFilename)
netIncome = fa.net_income(wb, wb['Worksheet'], exportFilename)

#cs = fa.customers_type(filename,wb)
#pm = fa.customers_payment_method(filename,wb)
#customersNumberByBranch = fa.branch_customers(filename,wb)
#categorySales = fa.category_sales(filename,wb)
#netIncome = fa.net_income(filename,wb)