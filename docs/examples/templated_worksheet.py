from os.path import dirname, join
from openpyxl_templates import TemplatedWorksheet, TemplatedWorkbook


class DictSheet(TemplatedWorksheet):
    def write(self, data):
        worksheet = self.worksheet

        for item in data.items():
            worksheet.append(list(item))

    def read(self):
        worksheet = self.worksheet
        data = {}

        for row in worksheet.rows:
            data[row[0].value] = row[1].value

        return data


class DictWorkbook(TemplatedWorkbook):
    dict_sheet = DictSheet(sheetname="dict_sheet")


workbook = DictWorkbook()

workbook.dict_sheet.write({
    "key1": "value1",
    "key2": "value2",
    "key3": "value3",
})

workbook.save("key_value_pairs.xlsx")

workbook = DictWorkbook(join(dirname(__file__), "key_value_pairs.xlsx"))

print(workbook.dict_sheet.read())
