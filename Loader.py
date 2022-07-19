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
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter
    from os.path import exists
    from plistlib import InvalidFileException

except ImportError as i:
    print(i.msg)

class Loader:
    '''
    Loads information pulled from documents into a new or existing Excel spreadsheet
    while also formatting the information and sheets

    Class attrs:

    att COLUMN_WIDTH: a constant specifying the width of the title columns
    inv: COLUMN_WIDTH is an int

    att COLUMN_HEADERS: a constant specifying the list of column title headers
    inv: COLUMN_HEADERS is a 1-D list of strings

    att reader: this Loader's associated Reader object
    inv: reader is an object of class Reader

    att excelname: the prefix of the Excel filename to be saved, 'newsheet' by default
    inv: excelname is a string

    att wstitle: the name of the saved spreadsheet sheet, 'Sheet' by default
    inv: wstitle is a string

    att wb: the active Workbook
    inv: wb is a Workbook or None
    '''
    # Hidden attrs:

    # att new: the name of the existing Excel sheet
    # inv: new is a string or None

    COLUMN_WIDTH = 22

    COLUMN_HEADERS = [
        'Agreement Type', 
        'Terminal', 
        'Customer Name - Corporate', 
        'Customer Name on Contract', 
        'Zenith Entity', 
        'Contract Date', 
        'Assignability/Transferrability',
        'Page Number'
        ]

    def setexcelname(self, input): # assert isalpha, prevent illegal characters
        '''
        Setter for excelname
        Param: input must be a string of Windows-legal characters
        '''
        self.excelname = input

    def setwstitle(self, input):
        '''
        Setter for wstitle
        Param: input must be a string
        '''
        self.wstitle = input

    def getreader(self):
        '''
        Returns associated Reader object
        '''
        return self.reader

    def getnew(self):
        '''
        Returns value of new attribute
        '''
        return self.new

    def setnew(self, input):
        '''
        Setter for new attribute
        Param: new must be a string or None
        '''
        self.new = input

    def __init__(self):
        '''
        Initializer
        '''
        self.new = None
        self.reader = Reader()
        self.excelname = 'newsheet' # default
        self.wstitle = 'Sheet' # default
        self.wb = None

    def excelinit(self):
        '''
        For new workbook, initializes a new spreadsheet and populates it with column titles
        '''
        if self.new == None:
            try:
                self.wb = Workbook()
            except OSError as e:
                print(e.errno)
            self.sheet = self.wb.active
            self.sheet.title = self.wstitle

            for header_number in range(1, len(self.COLUMN_HEADERS)+1):
                self.sheet.cell(column=header_number, row=1, value=self.COLUMN_HEADERS[header_number-1])
                letter = get_column_letter(header_number)
                self.sheet.column_dimensions[letter].width = self.COLUMN_WIDTH

        else:
            try:
                self.wb = load_workbook(filename = self.new)
                self.sheet = self.wb.active
            except InvalidFileException as e:
                print(e)

    def excelload(self):
        '''
        Calls reader's convert() method and loads this data into sheet within wb
        '''

        self.reader.convert()

        if self.reader.ret != []:
            for docdata in self.reader.ret:
                row_to_use = self.sheet.max_row + 1
                self.sheet.cell(column=1, row=row_to_use, value=docdata.agrtype()) # A: Agreement type

                self.sheet.cell(column=2, row=row_to_use, value=docdata.terminal()) # B: Terminal name

                self.sheet.cell(column=3, row=row_to_use, value=docdata.custname()) # C: Customer name on file

                self.sheet.cell(column=4, row=row_to_use, value=None) # D: Customer name on contract

                self.sheet.cell(column=5, row=row_to_use, value=docdata.entity()) # E: Company entity
                self.sheet['E{}'.format(str(row_to_use))].alignment = Alignment(wrapText=True)

                self.sheet.cell(column=6, row=row_to_use, value=None) # F: Contract date

                self.sheet.cell(column=7, row=row_to_use, value=docdata.userfield()) # G: Data to pull (assignability, etc.)
                self.sheet.column_dimensions['G'].width = 77

                self.sheet.cell(column=8, row=row_to_use, value=docdata.pagenumber()) # H: Page number of clause
                self.sheet['G{}'.format(str(row_to_use))].alignment = Alignment(wrapText=True)

    def savewb(self):
        '''
        If new == None, saves a new file, otherwise saves the existing Excel spreadsheet
        '''
        if self.new == None:
            dest = self.excelname + '.xlsx'
            if not exists(dest):
                self.wb.save(filename = dest)
            else:
                print('this file name already exists!')
        else:
            dest = self.new
            self.wb.save(filename = dest)
        self.clear()

    def clear(self):
        '''
        Clears Loader fields and calls associated Reader object's clear() method
        '''
        self.getreader().clear()
        self.new = None
        self.wb = None