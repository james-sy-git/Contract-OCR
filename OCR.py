'''
OCR Implementation using Tesseract and Open CV
'''

try:
    from PIL import Image
    import pytesseract
    import cv2
    import glob
    import fitz

    pytesseract.pytesseract.tesseract_cmd = 'C:\Program Files\Tesseract-OCR\\tesseract.exe'

except ImportError as i:
    print(i.msg)

class Reader:

    def __init__(self, filename):
        self.cache = ''
        self.convert()
        self.files = glob.glob(r'C:\\Users\\jsy13\Desktop\\OCRScr\\' + '*.png')
        self.process()
        try:
            print('Happy yay!')
        except OSError as e:
            print(e.errno)

    def process(self):

        for file in self.files:
            img = cv2.imread(file, cv2.IMREAD_GRAYSCALE)
            ship = Image.fromarray(img)
            final = ship.convert('RGB')
            add = pytesseract.image_to_string(final)

            self.cache = self.cache + add

    def convert(self):
        zoom_x = 7
        zoom_y = 7
        mat = fitz.Matrix(zoom_x, zoom_y)

        path = r'C:\\Users\\jsy13\Desktop\\OCRScr\\'
        all_files = glob.glob(path + '*.pdf')

        for filename in all_files:
            doc = fitz.open(filename)
            for page in doc:
                pix = page.get_pixmap(matrix = mat)
                pix.save(r'C:\Users\jsy13\Desktop\OCRScr\page-%i.png' % page.number)        

    def pull(self, input):
        pass

# if __name__ == '__main__':
#     test = Reader('voidedcontract.pdf')
#     out = test.cache
#     print('\n' in out)
#     print(out)
