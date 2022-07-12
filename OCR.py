'''
OCR Implementation using Tesseract and PyMuPDF
'''
try:
    from PIL import Image
    import pytesseract
    from cv2 import imread, IMREAD_GRAYSCALE
    from glob import glob
    from fitz import open, Matrix
    from os import path, remove
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
            self.newdoc = None
        except OSError as e:
            print(e.errno)

    def ask_conv(self):
        self.files_to_convert = askopenfilenames()

    def ask_dir(self):
        self.directory = askdirectory()

    def process(self):

        for file in self.files:

            if self.over != True:

                img = imread(file, IMREAD_GRAYSCALE)
                ship = Image.fromarray(img)
                final = ship.convert('RGB')
                add = pytesseract.image_to_string(final, config='--psm 4')

                if self.keyword in add:
                    self.newdoc.setuserfield(self.pull(add))
                    print('found it!')
                    self.over = True

            remove(file)

        self.ret.append(self.newdoc)
        self.newdoc = None

    def convert(self):
        zoom_x = 2
        zoom_y = 2
        mat = Matrix(zoom_x, zoom_y)

        for filename in self.files_to_convert:

            self.over = False

            self.newdoc = DocData(self.pullagrtype(filename))
            if self.newdoc.agrtype() != 'Multi':
                self.newdoc.setterminal(self.pullterminal(filename))
            self.newdoc.setcustomer(self.pullcustname(str(filename)))

            doc = open(filename) # double-check that this isn't Python built-in open
            for page in doc:
                pix = page.get_pixmap(matrix = mat)
                pix.save(self.directory + '\page-%i.png' % page.number)

            self.files = glob(self.directory + '\\' + '*.png')

            self.process()       

    def pull(self, document):
        interest = document.find(self.keyword)
        slice = document[interest:]
        parabreak = slice.find('\n\n')
        if parabreak != -1:
            chunk = slice[:parabreak]
            chunk.replace('\n', '')
            chunk.replace('\r', '')
            return chunk

    def pullagrtype(self, file):
        file_path = str(path.abspath(file))
        if 'Multi Terminal' in file_path:
            return 'Multi'
        else:
            return 'Single'
        
    def pullterminal(self, file):
        dir = path.dirname(file)
        return path.basename(dir)

    def pullcustname(self, file):
        stringed = str(path.basename(file))
        dash = stringed.find('-')
        return stringed[:dash]

    def pullentity(self, document):
        pass

    def clear(self):
        self.ret = []
        self.keyword = None

class DocData:

    def agrtype(self):
        return self._agreementtype

    def setagrtype(self, input):
        self._agreementtype = input

    def terminal(self):
        return self._terminal

    def setterminal(self, input):
        self._terminal = input

    def custname(self):
        return self._custname

    def setcustomer(self, input):
        self._custname = input

    def entity(self):
        return self._zentity

    def userfield(self):
        return self._pullfield

    def setuserfield(self, input):
        self._pullfield = input
    
    def __init__(self, agr):
        self._agreementtype = agr
        self._terminal = None
        self._custname = None
        self._zentity = None
        self._pullfield = None

# if __name__ == '__main__':
#     test = Reader()
#     test.setkeyword('assignability')
#     test.convert()
#     print(test.ret[0].userfield())
#     print(test.ret[0].agrtype())
#     print(test.ret[0].terminal())
#     print(test.ret[0].custname())
