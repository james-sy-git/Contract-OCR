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
        self.convert()
        self.file = 'page-0.png' # filename
        try:
            self.output = self.process()
        except OSError as e:
            print(e.errno)

    def process(self):
        return pytesseract.image_to_string(self.clarify())

    def clarify(self):
        # conv = self.convert(self.file)
        img = cv2.imread(self.file, cv2.IMREAD_GRAYSCALE)
        # denoised = cv2.fastNlMeansDenoising(img,None,10,7,21)
        # normalized = cv2.normalize(denoised, None, alpha=0, beta=3, norm_type=cv2.NORM_MINMAX, dtype=cv2.CV_32F)
        # gaussed = cv2.GaussianBlur(denoised, (1, 1), 0)
        final = Image.fromarray(img)

        return final.convert('RGB')

    def convert(self):
        zoom_x = 10
        zoom_y = 10
        mat = fitz.Matrix(zoom_x, zoom_y)

        path = r'C:\\Users\\jsy13\Desktop\\OCRScr\\'
        all_files = glob.glob(path + '*.pdf')

        for filename in all_files:
            doc = fitz.open(filename)
            for page in doc:
                pix = page.get_pixmap(matrix = mat)
                print("made matrix")
                pix.save(r'C:\Users\jsy13\Desktop\OCRScr\page-%i.png' % page.number)        

# if __name__ == '__main__':
#     test = Reader('voidedcontract.pdf')
#     out = test.process()
#     print('\n' in out)
