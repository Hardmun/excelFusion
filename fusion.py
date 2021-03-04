import os
import sys
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from xlwings import App


def runFusion():
    """global path"""
    projectDir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(
        os.path.abspath(__file__))

    excelList = [os.path.join(projectDir, "test_files", "yar.xlsx"),
                 os.path.join(projectDir, "test_files", "Client.xlsx")]

    mainFile = ""
    addFile = ""
    # for file in excelList:

    app = App(visible=False)
    excel1 = app.books.open(os.path.join(projectDir, "test_files", "yar.xlsx"))
    # excel2 = Book(os.path.join(projectDir, "test_files", "Челябинск.xlsx"))

    # app = App(visible=False)
    excelCommon = app.books.open(os.path.join(projectDir, "test_files", "Client.xlsx"))

    ws1 = excel1.sheets(1)
    ws1.name = "yar"
    ws1.api.Copy(Before=excelCommon.sheets(1).api)

    excelCommon.save()
    excelCommon.app.quit()


if __name__ == '__main__':
    runFusion()
