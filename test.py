import re


cleanString = re.sub('\W+'," ", "input string here" )

print(cleanString)

#This is a test 



filename = "example.xlsx"
wb = openpyxl.load_workbook(filename)
sheet = wb['Sheet1']
sheet.delete_rows(row_number, 1)
wb.save(filename)


COMBINED_RPT = "Report\Combined"


combined_wb = xw.Book()


for excel_file in excel_files:
    wb = xw.Book(excel_file)
    for sheet in wb.sheets:
        sheet.api.copy(After=combined_wb.sheets[0].api)
    wb.close()

combined_wb.sheets[0].delete()
combined_wb.save(f'all_worksheets_{timestamp}.xlsx')

if len(combined_wb.app.books) == 1:
    combined_wb.app.quit()
else:
    combined_wb.close()