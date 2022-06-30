'''
OCR Implementation using Tesseract and Open CV
'''

try:
    from PIL import Image
    import pytesseract
    import numpy as np
    import cv2
    pytesseract.pytesseract.tesseract_cmd = 'C:\Program Files\Tesseract-OCR\\tesseract.exe'

except:
    print('One or more libraries not found')

class Reader:

    def __init__(self, filename):
        self.file = filename
        self.output = None

    def process(self):
        return pytesseract.image_to_string(self.clarify())

    def clarify(self):
        img = cv2.imread(self.file, cv2.IMREAD_GRAYSCALE)
        denoised = cv2.fastNlMeansDenoising(img,None,50,7,21)
        normalized = cv2.normalize(denoised, None, alpha=0, beta=3, norm_type=cv2.NORM_MINMAX, dtype=cv2.CV_32F)
        gaussed = cv2.GaussianBlur(normalized, (1, 1), 0)
        final = Image.fromarray(gaussed)

        return final.convert('RGB')

if __name__ == '__main__':
    test = Reader('sample2.jpg')
    print(test.process())
