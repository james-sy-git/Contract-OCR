'''
Loads post-OCR PDFs into an Excel sheet using OpenPyXL
'''

try:
    from OCR import Reader
    from openpyxl import Workbook
    from openpyxl.styles import Alignment
except ImportError as i:
    print(i.msg)

class Loader:

    def setexcelname(self, input): # assert isalpha, prevent illegal characters
        self.excelname = input

    def __init__(self):
        self.reader = Reader()
        self.excelname = None
        self.wstitle = 'Sheet1'

    def excelload(self, input):
        wb = Workbook()
        dest = self.excelname + '.xlsx'

        ws1 = wb.active
        ws1.title = self.wstitle

        ws1.column_dimensions['B'].width = 150
        ws1['B2'].alignment = Alignment(wrapText=True)
        ws1.cell(column=2, row=2, value=input)
        wb.save(filename=dest)