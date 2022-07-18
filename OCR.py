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
                if self.newdoc.entity() == None:
                    self.newdoc.setentity(self.pullentity(add))

                if self.keyword in add:
                    self.newdoc.setuserfield(self.pull(add))
                    self.newdoc.setpagenumber(self.pullpage(file))
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

            doc = open(filename)
            for page in doc:
                pix = page.get_pixmap(matrix = mat)
                page.number = page.number + 1
                pix.save(self.directory + '\page-%i.png' % page.number)

            self.files = glob(self.directory + '\\' + '*.png')

            self.process()

    def isready(self):
        return (self.files_to_convert) != None and (self.directory) != None and (self.keyword) != ''     

    def pull(self, document):
        interest = document.find(self.keyword)
        slice = document[interest:]
        parabreak = slice.find('\n\n')
        if parabreak != -1:
            chunk = slice[:parabreak]
            chunk.replace('\n', '')
            chunk.replace('\r', '')
            chunk.strip()
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

    def pullentity(self, text): # could account for newline, but would severely affect runtime
        ents = ('Arc Terminals New York Holdings LLC',
        'Arc Terminals Holdings LLC',
        'Arc Terminals Pennsylvania Holdings LLC',
        'Zenith Energy Terminals Holdings LLC',
        'Zenith Energy Pennsylvania Holdings LLC',
        'Zenith Energy New York Holdings LLC')
        for ent in ents:
            if ent in text:
                return ent

    def pullpage(self, file):
        strfile = str(path.basename(file))
        rempage = strfile.replace('page-', '')
        return rempage.replace('.png', '')

    def clear(self):
        self.ret = []
        self.keyword = None
        self.files_to_convert = None
        self.directory = None

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

    def setentity(self, input):
        self._zentity = input

    def userfield(self):
        return self._pullfield

    def setuserfield(self, input):
        self._pullfield = input

    def pagenumber(self):
        return self._pagenumber

    def setpagenumber(self, input): # assert isdigit
        self._pagenumber = input

    def __init__(self, agr):
        self._agreementtype = agr
        self._terminal = None
        self._custname = None
        self._zentity = None
        self._pullfield = None
        self._pagenumber = None
