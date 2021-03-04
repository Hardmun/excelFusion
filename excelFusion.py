import os
import sys

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

def runFusion():
    """global path"""
    projectDir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(
        os.path.abspath(__file__))

    excel = load_workbook(os.path.join(projectDir, "test_files", "Ярославль.xlsx"))
    sheet = excel.active

    for row in sheet.iter_rows():
        for cell in row:
            if cell.comment:
                comment: str = cell.comment.text
                find_separator = comment.find("|")

                if find_separator != -1:
                    "cell format"
                    cell_Format = comment[:find_separator].replace("format_cell:", "")
                    cell.number_format = cell_Format
                    "formula"
                    formula = comment[-(len(comment) - find_separator - 1):].replace("formula_R1C1:", "")
                    cell.value = r1c1_to_a1(row=cell.row, column=cell.column, formula=formula).replace(";", ",")
                elif comment.find("formula_R1C1:") != -1:
                    "formula"
                    formula = comment.replace("formula_R1C1:", "")
                    cell.value = r1c1_to_a1(row=cell.row, column=cell.column, formula=formula).replace(";", ",")
                elif comment.find("format_cell:") != -1:
                    "format"
                    cell_Format = comment.replace("format_cell:", "")
                    cell.number_format = cell_Format

                    "converting string to float"
                    cellValue = cell.value
                    if cellValue is not None:
                        cell.value = float(cellValue.replace(",", ".").replace(" ", ""))

                else:
                    "formula"
                    cell.value = r1c1_to_a1(row=cell.row, column=cell.column, formula=comment).replace(";", ",")

    excel.save(os.path.join(projectDir, "test_files", "Ярославль_remastered.xlsx"))

if __name__ == '__main__':
    runFusion()
