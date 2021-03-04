import os
import sys

import win32com.client as win32

def insert_formulas(sheet):
    for comment in sheet.Comments:
        comment_text: str = comment.Text()
        find_separator = comment_text.find("|")
        "COM Object"
        cell_COM = comment.Parent

        if find_separator != -1:
            cell_Format = comment_text[:find_separator].replace("format_cell:", "")
            formula: str = comment_text[-(len(comment_text) - find_separator - 1):]

            cell_COM.NumberFormat = cell_Format
            cell_COM.NumberFormatLocal = cell_Format

            if formula.find("formula_R1C1:") != -1:
                cell_COM.FormulaR1C1Local = formula.replace("formula_R1C1:", "")
            else:
                cell_COM.FormulaLocal = formula
        elif comment_text.find("formula_R1C1:") != -1:
            cell_COM.FormulaR1C1Local = comment_text.replace("formula_R1C1:", "")
        elif comment_text.find("format_cell:") != -1:
            cell_COM.NumberFormatLocal = comment_text.replace("format_cell:", "")
        else:
            cell_COM.FormulaLocal = comment_text

def runFusion():
    """global path"""
    projectDir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(
        os.path.abspath(__file__))

    Client_Summary = os.path.join(projectDir, "test_files", "Client Summary .xlsx")
    yaroslavl = os.path.join(projectDir, "test_files", "Ярославль.xlsx")
    chelyabinsk = os.path.join(projectDir, "test_files", "Челябинск.xlsx")
    set = os.path.join(projectDir, "test_files", "сеть.xlsx")

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    workbook = excel.Workbooks.Add()  # Open(os.path.join(projectDir, "test_files", "common.xlsx"))

    # "1st list"
    # workbook_tmp = excel.Workbooks.Open(yaroslavl)
    # sheet_tmp = workbook_tmp.ActiveSheet
    # sheet_tmp.Name = "Ярославль"
    # sheet_tmp.Copy(workbook.Worksheets(1))
    # workbook_tmp.Close(False)
    # sheet = workbook.Worksheets(1)
    # insert_formulas(sheet)

    # "2nd list"
    # workbook_tmp = excel.Workbooks.Open(chelyabinsk)
    # sheet_tmp = workbook_tmp.ActiveSheet
    # sheet_tmp.Name = "Челябинск"
    # sheet_tmp.Copy(workbook.Worksheets(1))
    # workbook_tmp.Close(False)
    # sheet = workbook.Worksheets(1)
    # insert_formulas(sheet)
    #
    # "3nd list"
    # workbook_tmp = excel.Workbooks.Open(set)
    # sheet_tmp = workbook_tmp.ActiveSheet
    # sheet_tmp.Name = "сеть"
    # sheet_tmp.Copy(workbook.Worksheets(1))
    # workbook_tmp.Close(False)
    # sheet = workbook.Worksheets(1)
    # insert_formulas(sheet)
    #
    "4th list"
    workbook_tmp = excel.Workbooks.Open(Client_Summary)
    sheet_tmp = workbook_tmp.ActiveSheet
    sheet_tmp.Name = "Client_Summary"
    sheet_tmp.Copy(workbook.Worksheets(1))
    workbook_tmp.Close(False)
    sheet = workbook.Worksheets(1)
    insert_formulas(sheet)

    workbook.SaveAs(os.path.join(projectDir, "test_files", "common.xlsx"))
    excel.Application.Quit()

if __name__ == '__main__':
    runFusion()
