import os
import sys

import win32com.client as win32

def runFusion():
    """global path"""
    projectDir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(
        os.path.abspath(__file__))

    os.path.join(projectDir, "test_files", "ЯрославльCOM.xlsx")
    # excel = win32.gencache.EnsureDispatch('Excel.Application')
    #
    # excel.Application.Quit()

if __name__ == '__main__':
    runFusion()
