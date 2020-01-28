import openpyxl
from openpyxl.utils.exceptions import *

if __name__ == '__main__':
    w = None
    try:
        w = openpyxl.load_workbook('../Template.xlsx')
    except InvalidFileException as e:
        print(e.args)
    w.save