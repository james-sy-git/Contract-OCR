'''
Loads post-OCR PDFs into an Excel sheet using OpenPyXL
Buttons:
- enter keyword
- choose files
- choose working directory
- convert and search then load into Excel
'''

try:
    from OCR import Reader
    from openpyxl import Workbook
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter

except ImportError as i:
    print(i.msg)

class Loader:

    COLUMN_WIDTH = 22

    COLUMN_HEADERS = [
        'Agreement Type', 
        'Terminal', 
        'Customer Name - Corporate', 
        'Customer Name on Contract', 
        'Zenith Entity', 
        'Contract Date', 
        'Assignability/Transferrability'
        ]

    def setexcelname(self, input): # assert isalpha, prevent illegal characters
        self.excelname = input

    def setwstitle(self, input):
        self.wstitle = input

    def getreader(self):
        return self.reader

    def __init__(self):
        self.reader = Reader()
        self.reader.setkeyword('assignability')
        self.reader.convert()
        self.excelname = None
        self.wstitle = 'Sheet1'
        try:
            self.wb = Workbook()
        except OSError as e:
            print(e.errno)

    def excelinit(self):
        self.sheet = self.wb.active
        self.sheet.title = self.wstitle

        for header_number in range(1, len(self.COLUMN_HEADERS)+1):
            self.sheet.cell(column=header_number, row=1, value=self.COLUMN_HEADERS[header_number-1])
            letter = get_column_letter(header_number)
            self.sheet.column_dimensions[letter].width = self.COLUMN_WIDTH

    def excelload(self):

        if self.reader.ret != []:
            for docdata in self.reader.ret:
                row_to_use = self.sheet.max_row + 1
                self.sheet.cell(column=1, row=row_to_use, value=docdata.agrtype()) # A: Agreement type

                self.sheet.cell(column=2, row=row_to_use, value=docdata.terminal()) # B: Terminal name

                self.sheet.cell(column=3, row=row_to_use, value=docdata.custname()) # C: Customer name on file

                self.sheet.cell(column=4, row=row_to_use, value=None) # D: Customer name on contract

                self.sheet.cell(column=5, row=row_to_use, value=docdata.entity()) # E: Company entity

                self.sheet.cell(column=6, row=row_to_use, value=None) # F: Contract date

                self.sheet.cell(column=7, row=row_to_use, value=docdata.userfield()) # G: Data to pull (assignability, etc.)
                self.sheet.column_dimensions['G'].width = 77


        self.sheet['G2'].alignment = Alignment(wrapText=True)

    def savewb(self):
        dest = self.excelname + '.xlsx'
        self.wb.save(filename = dest)

if __name__ == "__main__":
    test = Loader()
    test.setexcelname('example2')
    test.excelinit()
    test.excelload()
    test.savewb()