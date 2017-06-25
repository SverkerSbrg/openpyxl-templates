from openpyxl_templates import TemplatedWorkbook, TemplatedWorksheet


class DemoTemplatedWorkbook(TemplatedWorkbook):
    sheet1 = TemplatedWorksheet()
    sheet2 = TemplatedWorksheet()


templated_workbook = DemoTemplatedWorkbook()

templated_workbook = DemoTemplatedWorkbook(filename="my_excel.xlsx")


for templated_worksheet in templated_workbook.templated_sheets:
    print(templated_worksheet.sheetname)


templated_workbook.save("my_excel.xlsx")