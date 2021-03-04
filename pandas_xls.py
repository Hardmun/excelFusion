import pandas as pd
import os, sys

def runFusion():
    """global path"""
    projectDir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(
        os.path.abspath(__file__))
    excel = pd.ExcelWriter(os.path.join(projectDir, "test_files", "test.xlsx"), engine="xlsxwriter")
    df = pd.DataFrame





if __name__ == '__main__':
    runFusion()