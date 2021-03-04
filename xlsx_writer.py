from xlsxwriter.workbook import Workbook
import os, sys

def runFusion():
    """global path"""
    projectDir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(
        os.path.abspath(__file__))

    excel = Workbook(os.path.join(projectDir, "test_files", "test.xlsx"))
    sheet = excel.get_worksheet_by_name("TDSheet")

if __name__ == '__main__':
    runFusion()
