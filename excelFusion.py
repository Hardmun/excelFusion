import os
import sys
from copy import copy

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def IsItNumber(value_check):
    try:
        conv_value = int(value_check)
        return isinstance(conv_value, int)
    except:
        return False

def r1c1_to_a1(row: int, column: int, formula: str):
    len_formula = len(formula)
    a1 = ""
    parsing_started = False
    start_Row = False
    currentRow = ""
    start_Col = False
    currentCol = ""

    for i, t in zip(formula, range(0, len(formula))):
        "rows parsing"
        if i.upper() == "R":
            next_symb = formula[t + 1]
            if next_symb.upper() == "C":
                currentRow = row
                parsing_started = True
            elif next_symb == "[" or IsItNumber(next_symb):
                start_Row = True
                parsing_started = True
        elif i == "[" and start_Row:
            pass
        elif i == "]" and start_Row:
            currentRow = row + int(currentRow)
            start_Row = False
        elif i.upper() == "C" and start_Row:
            currentRow = int(currentRow)
            start_Row = False
        elif start_Row:
            currentRow += i

        "column parsing"
        if i.upper() == "C":
            if (t + 1) < len_formula:
                next_symb = formula[t + 1]
                if next_symb == "[" or IsItNumber(next_symb):
                    start_Col = True
                elif not next_symb.isalpha():
                    currentCol = column
            else:
                currentCol = column
        elif i == "[" and start_Col:
            pass
        elif i == "]" and start_Col:
            currentCol = column + int(currentCol)
            start_Col = False
        elif start_Col:
            if (t + 1) < len_formula:
                next_symb = formula[t + 1]
                if not IsItNumber(next_symb):
                    currentCol += i
                    if next_symb != "]":
                        currentCol = int(currentCol)
                        start_Col = False
                else:
                    currentCol += i
            else:
                currentCol += i
                currentCol = int(currentCol)
                start_Col = False

        if not parsing_started:
            a1 += i
        elif isinstance(currentRow, int) and isinstance(currentCol, int):
            a1 += f"{get_column_letter(currentCol)}" \
                  f"{str(currentRow)}"
            currentRow = ""
            currentCol = ""
            parsing_started = False

    return a1

def copySheet(target, source):
    for (row, col), source_cell in source._cells.items():
        target_cell = target.cell(column=col, row=row)
        target_cell._value = source_cell._value
        target_cell.data_type = source_cell.data_type

        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)

        if source_cell.hyperlink:
            target_cell._hyperlink = copy(source_cell.hyperlink)

        if source_cell.comment:
            target_cell.comment = copy(source_cell.comment)

    for attr in ('row_dimensions', 'column_dimensions'):
        src = getattr(source, attr)
        trg = getattr(target, attr)
        for key, dim in src.items():
            trg[key] = copy(dim)
            trg[key].worksheet = trg

    target.sheet_format = copy(source.sheet_format)
    target.sheet_properties = copy(source.sheet_properties)
    target.merged_cells = copy(source.merged_cells)
    target.page_margins = copy(source.page_margins)
    target.page_setup = copy(source.page_setup)
    target.print_options = copy(source.print_options)

def insertFormulas(sheet):
    for (row, col), cell in sheet._cells.items():
        if cell.comment:
            comment: str = cell.comment.text
            find_separator = comment.find("|")
            deleteComment = True

            if find_separator != -1:
                "cell format"
                cell_Format = comment[:find_separator].replace("format_cell:", "")
                cell.number_format = cell_Format
                "formula"
                formula = comment[-(len(comment) - find_separator - 1):].replace("formula_R1C1:", "")
                cell.value = r1c1_to_a1(row=row, column=col, formula=formula).replace(";", ",")
            elif comment.find("formula_R1C1:") != -1:
                "formula"
                formula = comment.replace("formula_R1C1:", "")
                cell.value = r1c1_to_a1(row=row, column=col, formula=formula).replace(";", ",")
            elif comment.find("format_cell:") != -1:
                "format"
                cell_Format = comment.replace("format_cell:", "")
                cell.number_format = cell_Format

                "converting string to float"
                cellValue = cell.value
                if cellValue is not None:
                    cell.value = float(cellValue.replace(",", ".").replace(" ", ""))

            elif comment.find("formula:") != -1:
                "formula"
                formula = comment.replace("formula:", "")
                cell.value = r1c1_to_a1(row=row, column=col, formula=formula).replace(";", ",")
            else:
                deleteComment = False

            if deleteComment:
                cell.comment = None

def ExcelFusion(curr_file, fileExcel):
    wb_path = curr_file.get("file")
    ws_title = curr_file.get("title")

    """file exists"""
    if not os.path.isfile(wb_path):
        return f"File: {wb_path} doesn't exist!!!"

    if fileExcel is None:
        fileExcel = load_workbook(wb_path)
        ws = fileExcel.worksheets[0]
        ws.title = ws_title
        insertFormulas(ws)
    else:
        wb_from = load_workbook(wb_path)
        ws_from = wb_from.worksheets[0]
        insertFormulas(ws_from)
        ws = fileExcel.create_sheet(title=ws_title, index=0)
        copySheet(ws, ws_from)

    return fileExcel

def readFiles(fileSettings_string):
    """global path"""
    projectDir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(
        os.path.abspath(__file__))

    """convert to dict"""
    fileSettings = eval(fileSettings_string)

    """checking key - files"""
    settings = fileSettings.get("settings")

    files = settings.get("files")
    if files is not None:
        wb = None
        for curr_file in files:
            if curr_file:
                wb = ExcelFusion(curr_file, wb)

        saveAs = os.path.abspath(settings.get("SaveAs"))
        wb.save(saveAs)

if __name__ == '__main__':
    if len(sys.argv) == 2:
        files_to_read = sys.argv[1]
    else:
        raise Exception("Wrong parameters.")
    readFiles(files_to_read)
