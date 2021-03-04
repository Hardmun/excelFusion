import os
import sys
from copy import copy

from openpyxl import load_workbook, Workbook

class xls(Workbook):
    pass

"""global path"""
projectDir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(
    os.path.abspath(__file__))

def combineFiles():
    wb1 = load_workbook(os.path.join(projectDir, "test_files", "Client.xlsx"))
    ws1 = wb1.create_sheet("lol")

    wb2 = load_workbook(os.path.join(projectDir, "test_files", "Ярославль_remastered.xlsx"))
    ws2 = wb2.active
    for row in ws2.iter_rows():
        for cell in row:
            try:
                new_cell = ws1.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = copy(cell.number_format)
                    new_cell.protection = copy(cell.protection)
                    new_cell.alignment = copy(cell.alignment)
            except:
                asdf = 0
            if cell.has_style:
                pass
                # new_cell.font = cell.font
                # new_cell.border = cell.border
                # new_cell.fill = cell.fill
                # new_cell.number_format = cell.number_format
                # new_cell.protection = cell.protection
                # new_cell.alignment = cell.alignment

                # new_cell._style = copy(cell._style)
        # ws1.append((cell.value for cell in row))

    wb1.save(os.path.join(projectDir, "test_files", "fusion.xlsx"))

def combineFilesByopenpyxl():
    wb1 = Workbook()
    # wb1 = load_workbook(os.path.join(projectDir, "test_files", "Client.xlsx"))

    wb2 = load_workbook(os.path.join(projectDir, "test_files", "Ярославль_remastered.xlsx"))
    ws2 = wb2.active

    # ws1 = wb1.copy_worksheet(ws2)

    wb1.save(os.path.join(projectDir, "test_files", "fusion.xlsx"))

def copyCell():
    wb1 = load_workbook(os.path.join(projectDir, "test_files", "Client.xlsx"))
    ws1 = wb1.create_sheet("lol")

    wb2 = load_workbook(os.path.join(projectDir, "test_files", "Ярославль_remastered.xlsx"))
    ws2 = wb2.active
    for row in ws2.iter_rows():
        for cell in row:
            try:
                new_cell = ws1.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = copy(cell.number_format)
                    new_cell.protection = copy(cell.protection)
                    new_cell.alignment = copy(cell.alignment)
            except:
                asdf = 0
            if cell.has_style:
                pass
                # new_cell.font = cell.font
                # new_cell.border = cell.border
                # new_cell.fill = cell.fill
                # new_cell.number_format = cell.number_format
                # new_cell.protection = cell.protection
                # new_cell.alignment = cell.alignment

                # new_cell._style = copy(cell._style)
        # ws1.append((cell.value for cell in row))

    wb1.save(os.path.join(projectDir, "test_files", "fusion.xlsx"))

def copySheet():
    wb1 = load_workbook(os.path.join(projectDir, "test_files", "Ярославль_remastered.xlsx"))
    sheetToCopy = wb1["TDSheet"]
    wb1.copy_worksheet(sheetToCopy)

    # wb2 = load_workbook(os.path.join(projectDir, "test_files", "Ярославль_remastered.xlsx"))
    # ws2 = wb2.active
    wb1.save(os.path.join(projectDir, "test_files", "fusion.xlsx"))

if __name__ == '__main__':
    copySheet()
    # copyCell()
    # combineFilesByopenpyxl()
    # combineFiles()
