'''
OCR Implementation using Tesseract and Open CV
'''
try:
    from PIL import Image
    import pytesseract
    import cv2
    import glob
    import fitz
    import os
    from openpyxl import Workbook
    from openpyxl.styles import Alignment

    pytesseract.pytesseract.tesseract_cmd = 'C:\Program Files\Tesseract-OCR\\tesseract.exe'

except ImportError as i:
    print(i.msg)

class Reader:

    def __init__(self, keyword):
        try:
            self.keyword = keyword.upper()
            self.ret = None
            self.over = False
            self.convert()
            self.files = glob.glob(r'C:\\Users\\jsy13\Desktop\\OCRScr\\' + '*.png')
            self.process()
        except OSError as e:
            print(e.errno)

    def process(self):

        for file in self.files:

            if self.over != True:

                img = cv2.imread(file, cv2.IMREAD_GRAYSCALE)
                ship = Image.fromarray(img)
                final = ship.convert('RGB')
                add = pytesseract.image_to_string(final, config='--psm 4')

                if self.keyword in add:
                    self.ret = self.pull(add)
                    print('found it!')
                    self.over = True
                    self.excelload(self.ret)
                    # print(self.ret)

            os.remove(file)

    def convert(self):
        zoom_x = 2
        zoom_y = 2
        mat = fitz.Matrix(zoom_x, zoom_y)

        path = r'C:\\Users\\jsy13\Desktop\\OCRScr\\'
        all_files = glob.glob(path + '*.pdf')

        for filename in all_files:
            doc = fitz.open(filename)
            for page in doc:
                pix = page.get_pixmap(matrix = mat)
                pix.save(r'C:\Users\jsy13\Desktop\OCRScr\page-%i.png' % page.number)        

    def pull(self, document):
        interest = document.find(self.keyword)
        slice = document[interest:]
        parabreak = slice.find('\n\n')
        if parabreak != -1:
            chunk = slice[:parabreak]
            return chunk

    def excelload(self, input):
        wb = Workbook()
        dest = 'example.xlsx'

        ws1 = wb.active
        ws1.title = 'Assignability Clauses'

        ws1.column_dimensions['B'].width = 150
        ws1['B2'].alignment = Alignment(wrapText=True)
        ws1.cell(column=2, row=2, value=input)
        wb.save(filename=dest)

if __name__ == '__main__':
    test = Reader('assignability')
