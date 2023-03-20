import fileAutomation as fa

# call loading func
filename = 'sales.xlsx'
wb = fa.load_workbook(filename)

# call data automation funcions

paymentsByGender = fa.payment_method_by_gender(filename,wb)

#cs = fa.customers_type(filename,wb)
#pm = fa.customers_payment_method(filename,wb)
#customersNumberByBranch = fa.branch_customers(filename,wb)
#categorySales = fa.category_sales(filename,wb)
#netIncome = fa.net_income(filename,wb)