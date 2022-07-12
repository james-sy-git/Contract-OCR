'''
OCR Implementation using Tesseract and PyMuPDF
'''
try:
    from PIL import Image
    import pytesseract
    import cv2
    import glob
    import fitz
    import os
    # from openpyxl import Workbook
    # from openpyxl.styles import Alignment
    from tkinter.filedialog import askopenfilenames, askdirectory
    from tkinter import Tk

    Tk().wm_withdraw()

    pytesseract.pytesseract.tesseract_cmd = 'C:\Program Files\Tesseract-OCR\\tesseract.exe'

except ImportError as i:
    print(i.msg)

class Reader:

    def setkeyword(self, keyword):
        self.keyword = keyword.upper()

    def __init__(self):
        try:
            self.keyword = None
            self.files_to_convert = None
            self.directory = None
            self.ask_conv()
            self.ask_dir()
            self.ret = []
            self.over = False
            self.files = None
        except OSError as e:
            print(e.errno)

    def ask_conv(self):
        self.files_to_convert = askopenfilenames()

    def ask_dir(self):
        self.directory = askdirectory()

    def process(self):

        for file in self.files:

            if self.over != True:

                img = cv2.imread(file, cv2.IMREAD_GRAYSCALE)
                ship = Image.fromarray(img)
                final = ship.convert('RGB')
                add = pytesseract.image_to_string(final, config='--psm 4')

                if self.keyword in add:
                    self.ret.append(self.pull(add))
                    print('found it!')
                    self.over = True
                    print(self.ret)

            os.remove(file)

    def convert(self):
        zoom_x = 2
        zoom_y = 2
        mat = fitz.Matrix(zoom_x, zoom_y)

        for filename in self.files_to_convert:

            self.over = False

            doc = fitz.open(filename)
            for page in doc:
                pix = page.get_pixmap(matrix = mat)
                pix.save(self.directory + '\page-%i.png' % page.number)

            self.files = glob.glob(self.directory + '\\' + '*.png')

            self.process()       

    def pull(self, document):
        interest = document.find(self.keyword)
        slice = document[interest:]
        parabreak = slice.find('\n\n')
        if parabreak != -1:
            chunk = slice[:parabreak]
            return chunk

    def clear(self):
        self.ret = []
        self.keyword = None

if __name__ == '__main__':
    test = Reader()
    test.setkeyword('assignability')
    test.convert()
