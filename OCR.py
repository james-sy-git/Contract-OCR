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

    # Tk().wm_withdraw() # headless Tkinter GUI to use file dialog

    pytesseract.pytesseract.tesseract_cmd = 'C:\Program Files\Tesseract-OCR\\tesseract.exe'

except ImportError as i:
    print(i.msg)

class Reader:
    '''
    Converts PDFs into sequenced PNGs and uses Tesseract OCR to extract text.
    Pulls out relevant information from this text and stores them in a DocData
    object. Searches the document for a clause matching the user-inputted keyword.

    Class attrs:

    att keyword: the user-inputted keyword
    inv: keyword is a string or None

    att files_to_convert: the PDF files to be converted into PNGs and then processed
    inv: files_to_convert is a list of PDF files or None

    att directory: the working directory in which the conversion occurs and the output is given
    inv: directory is an existing directory somewhere on the user's computer

    att ret: the list of DocData objects containing document information
    inv: ret is a list of DocData objects or an empty list

    att over: whether or not the processing loop should continue to search pages
    inv: over is a boolean

    att files: the list of PNG files in the directory to be OCR'd and parsed
    inv: files is a list of PNG files or None

    att newdoc: the DocData object associated with a PDF
    inv: newdoc is a DocData object or None
    '''

    def setkeyword(self, keyword):
        '''
        Setter for keyword that turns the input into uppercase
        Param: keyword must be a string
        '''
        self.keyword = keyword.upper()

    def setfiles(self, files):
        self.files_to_convert = files

    def setdir(self, directory):
        self.directory = directory

    def __init__(self):
        '''
        Initializer
        '''
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

    def process(self):
        '''
        Uses Tesseract to extract text from the PNGs in files, then searches
        for the keyword and entity. Also fills the page number where the keyword
        clause was found. Removes the just-searched file after every iteration, and
        stops the loop after the keyword clause was found. Adds each new DocData
        object to ret.
        '''

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
        '''
        Converts each file of files_to_convert from a PDF to a sequence of PNGs
        by page. Zooms in to achieve better resolution. Pulls the terminal name,
        agreement type, and customer from the filepath. Saves the PNG images to
        the specified directory.
        '''
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
        '''
        Returns whether or nto the Reader is ready to accept and process documents.
        '''
        return (self.files_to_convert) != None and (self.directory) != None and (self.keyword) != '' and (self.keyword) != None     

    def pull(self, document):
        '''
        Searches for the keyword clause, returns it if found
        Param: document is a string
        '''
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
        '''
        Extracts and returns the agreement type
        Param: file is a valid, existing file
        '''
        file_path = str(path.abspath(file))
        if 'Multi Terminal' in file_path:
            return 'Multi'
        else:
            return 'Single'
        
    def pullterminal(self, file):
        '''
        Extracts and returns the terminal name
        Param: file is a valid, existing file
        '''
        dir = path.dirname(file)
        return path.basename(dir)

    def pullcustname(self, file):
        '''
        Extracts and returns the customer name
        Param: file is a valid, existing file
        '''
        stringed = str(path.basename(file))
        dash = stringed.find('-')
        return stringed[:dash]

    def pullentity(self, text): # could account for newline, but would severely affect runtime
        '''
        Extracts and returns the company entity, if it is one of
        the names enumerated in ents.
        Param: text is a string
        '''
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
        '''
        Isolates the page number according to the file-naming conventions used
        by the convert() method
        Param: file is a valid, existing PNG file created by the convert() method
        '''
        strfile = str(path.basename(file))
        rempage = strfile.replace('page-', '')
        return rempage.replace('.png', '')

    def clear(self):
        '''
        Clear method to restore to defaults
        '''
        self.ret = []
        self.keyword = None
        self.files_to_convert = None
        self.directory = None

class DocData:
    '''
    Stores data that is extracted with a Reader
    '''

    def agrtype(self):
        '''
        Returns agreement type
        '''
        return self._agreementtype

    def setagrtype(self, input):
        '''
        Sets agreement type
        Param: input is a string, either 'Multi' or 'Single'
        '''
        self._agreementtype = input

    def terminal(self):
        '''
        Returns terminal name
        '''
        return self._terminal

    def setterminal(self, input):
        '''
        Sets terminal name
        Param: input is a string
        '''
        self._terminal = input

    def custname(self):
        '''
        Returns customer name
        '''
        return self._custname

    def setcustomer(self, input):
        '''
        Sets customer name
        Param: input is a string
        '''
        self._custname = input

    def entity(self):
        '''
        Returns company ownership entity name
        '''
        return self._zentity

    def setentity(self, input):
        '''
        Sets entity name
        Param: input is a string, one of ents in Loader's pullentity() method
        '''
        self._zentity = input

    def userfield(self):
        '''
        Returns user's keyword-searched field
        '''
        return self._pullfield

    def setuserfield(self, input):
        '''
        Sets user's keyword-searched field
        Param: input is a string
        '''
        self._pullfield = input

    def pagenumber(self):
        '''
        Returns page number
        '''
        return self._pagenumber

    def setpagenumber(self, input): # assert isdigit
        '''
        Sets page number
        Param: input is a string-converted int
        '''
        self._pagenumber = input

    def __init__(self, agr):
        '''
        Initializer
        Param agr: the agreement type, must be 'Multi' or 'Single'
        '''
        self._agreementtype = agr
        self._terminal = None
        self._custname = None
        self._zentity = None
        self._pullfield = None
        self._pagenumber = None
