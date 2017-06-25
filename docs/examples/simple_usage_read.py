from docs.examples.simple_usage_write import PersonSheet
from openpyxl_templates import TemplatedWorkbook


class PersonsWorkbook(TemplatedWorkbook):
    persons = PersonSheet()

wb = PersonsWorkbook("fruit_lovers.xlsx")

for person in wb.persons:
    print(person.first_name, person.last_name, person.favorite_fruit)

