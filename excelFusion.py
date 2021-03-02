import sys
import os
from openpyxl import load_workbook

def runFusion():
    """global path"""
    projectDir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(
        os.path.abspath(__file__))

    excel = load_workbook(os.path.join(projectDir, "test_files", "Ярославль.xlsx"))
    sheet = excel.active

    excel.save(os.path.join(projectDir, "test_files", "Ярославль_remastered.xlsx"))
    # with open(os.path.join(projectDir, "lollol.txt"),'w') as myfile:
    #     myfile.write("")
    #
    # for i in range(1, len(sys.argv)):
    #     print(argv[i])


if __name__ == '__main__':
    runFusion()
