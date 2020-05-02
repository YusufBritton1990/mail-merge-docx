from openpyxl import load_workbook #used to import the Excel Data
from mailmerge import MailMerge

wb = load_workbook('SampleData.xlsx') #open excel workbook
sheet = wb['SalesOrders'] #Tab to get information
max_row = sheet.max_row #count of all of the rows

# for i in range(1, max_row):
for i in range(2, 4): #just interating 3 for testing
    # template_doc = "OrderTemplate.docx" #word dock
    template_doc = "OrderTemplate.docx" #word dock
    word_doc = MailMerge(template_doc)

    word_doc.merge(
        Name = str(sheet.cell(row = i, column = 3).value),
        Date = str(sheet.cell(row = i, column = 1).value),
        Quantity = str(sheet.cell(row = i, column = 5).value),
        Cost = str(sheet.cell(row = i, column = 6).value),
        Subtotal = str(sheet.cell(row = i, column = 7).value)
        )

    word_doc.write("Invoice for " + str(sheet.cell(row = i, column = 3).value) + ".docx")
