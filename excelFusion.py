import sys
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# from openpyxl.comments import Comment

def IsItNumber(value_check):
    try:
        conv_value = int(value_check)
        return isinstance(conv_value, int)
    except:
        return False

def r1c1_to_a1(row: int, column: int, formula: str):
    a1 = ""
    txt_replace = ""
    start_Row = False
    currentRow = ""
    start_Col = False
    currentCol = ""
    replaceStarted = False
    for i, t in zip(formula, range(0, len(formula))):
        if i.upper() == "R" and formula[t + 1].upper() != "C":
            replaceStarted = True
            start_Row = True
        elif i == "[":
            pass
        elif (i == "]" or i.upper() == "C") and start_Row:
            start_Row = False
            if currentRow:
                currentRow = row + int(currentRow)
        elif start_Row:
            currentRow += i
        # elif i.upper() == "C" and t == len(formula):
        #     currentCol = column
        # elif i.upper() == "C" and not IsItNumber(formula[t+1]) and formula[t+1].upper() != "O":
        #     currentCol = column
        elif i.upper() == "C" and t < len(formula) and (IsItNumber(formula[t + 1]) or formula[t + 1] == "["):
            start_Col = True
        elif i == "[":
            pass
        elif i == "]" and start_Col:
            start_Col = False
            if currentCol:
                currentCol = column + int(currentCol)
            if i == ":":
                a1 += i
        elif start_Col:
            currentCol += i
        else:
            if replaceStarted:
                replaceStarted = False
                a1 += f"{get_column_letter(currentCol if currentCol else column)}" \
                      f"{str(currentRow) if str(currentRow) else row}"
                currentRow = ""
                currentCol = ""
            a1 += i

    return a1

    # print(f"{i}  {t}")

def runFusion():
    """global path"""
    projectDir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(
        os.path.abspath(__file__))

    # excel = load_workbook(os.path.join(projectDir, "test_files", "Ярославль.xlsx"))
    # sheet = excel.active

    # sheet.cell(1,"C").value = "=D1"
    # sheet.cell(row=39, column=9, value="=COUNT(R[-24]C:R[-1]C)")

    curForm = r1c1_to_a1(row=39, column=9, formula="=COUNT(R[-24]C:R[-1]C)")
    sdf = 0
    # sheet.cell(row=15, column=40, value="=COUNT(I15:I38)")

    # for row in sheet.iter_rows():
    #     for cell in row:
    #         if cell.comment:
    #             comment: str = cell.comment.text
    #             find_separator = comment.find("|")
    #
    #             if find_separator != -1:
    #                 cell_Format = comment[:find_separator].replace("format_cell:", "")
    #                 formula = comment[-(len(comment) - find_separator - 1):]
    #                 cell.value = formula
    #                 break
    #     else:
    #         continue
    #
    #     break

    # sheet
    # cmt = Comment(cell.comment.text)
    # sheet
    # print(cell.comment.text)

    # excel.save(os.path.join(projectDir, "test_files", "Ярославль_remastered.xlsx"))
    # with open(os.path.join(projectDir, "lollol.txt"),'w') as myfile:
    #     myfile.write("")
    #
    # for i in range(1, len(sys.argv)):
    #     print(argv[i])

if __name__ == '__main__':
    runFusion()
