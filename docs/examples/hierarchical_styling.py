from itertools import chain

from openpyxl_templates import TemplatedWorkbook
from openpyxl_templates.styles import ExtendedStyle
from openpyxl_templates.table_sheet import TableSheet, WriteOnlyCell
from openpyxl_templates.table_sheet.columns import CharColumn, IntColumn, RowStyle
from openpyxl_templates.utils import SolidFill


class Employee:
    def __init__(self, name, title, unit=None, manager=None, employees=None):
        self.name = name
        self.title = title
        self.manager = None
        self._unit = unit
        if manager:
            self.add_manager(manager)

        self.employees = []
        if employees:
            self.add_employees(*employees)

    def add_employees(self, *employees):
        for employee in employees:
            self.employees.append(employee)
            employee.manager = self

    def add_manager(self, manager):
        manager.add_employees(self)

    @property
    def level(self):
        if not self.manager:
            return 1

        return self.manager.level + 1

    @property
    def unit(self):
        if self._unit:
            return self._unit

        if not self.manager:
            return ""

        return str(self.manager.unit)

    @property
    def descendants(self):
        yield self
        for employee in self.employees:
            for descendant in employee.descendants:
                yield descendant

CEO = Employee(
    "Breanna Forster", "CEO",
    unit="Pear Company",
    employees=[
        Employee(
            "Jawad Guthrie", "CFO",
            unit="Financial services",
            employees=[
                Employee("Sylvia Payne", "Accountant"),
                Employee("Giorgio Edge", "Accountant")
            ]
        ),
        Employee(
            "Edgar Ortiz", "HR manager",
            unit="HR"
        ),
        Employee(
            "Rumaisa Albert", "Head of operations",
            unit="Operations",
            employees=[
                Employee("Wade Ortega", "Specialist"),
                Employee("Hereem Hope", "Specialist", employees=[Employee("Francesca Cabrera", "Trainee")]),
                Employee("Jerome Mills", "Specialist"),

            ]
        ),

        Employee(
            "Inez Watts", "Sales manager",
            unit="Sales",
            employees=[
                Employee("Ellie-Louise Mcnamara", "Specialist"),
                Employee("Sulaiman Cross", "Specialist")

            ]
        )

    ]
)


class OrgStructureWorksheet(TableSheet):
    name = CharColumn(header="Name")
    title = CharColumn(header="Title")
    unit = CharColumn(header="Unit")
    manager = CharColumn(header="Manager", getter=lambda column, obj: obj.manager.name if obj.manager else "")
    level = IntColumn(header="Level")

    hide_excess_columns = False

    row_styles = [
        RowStyle(
            row_type=1,
            cell_style=ExtendedStyle(None, lambda x: x + " level 1", fill=SolidFill("bbbbbb"))
        ),
        RowStyle(
            row_type=2,
            cell_style=ExtendedStyle(None, lambda x: x + " level 2", fill=SolidFill("dddddd"))

        ),
    ]

    def row_type(self, object, row_number):
        return object.level


class PearCompanyOrgWorkbook(TemplatedWorkbook):
    org_structure = OrgStructureWorksheet(sheetname="Org structure")


wb = PearCompanyOrgWorkbook()
wb.org_structure.write(list(CEO.descendants), title="Pear company organizational structure")

wb.save("pear_company.xlsx")