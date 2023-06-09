import openpyxl
import win32com.client
from managerbom import ManagerBOM


def main():
    template = openpyxl.load_workbook("BOM_template.xlsx")
    iam = win32com.client.Dispatch("Inventor.Application").ActiveDocument
    manager = ManagerBOM(template, iam)
    manager.issue_bom().save("BOM.xlsx")


if __name__ == "__main__":
    main()
