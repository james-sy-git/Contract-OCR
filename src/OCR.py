'''
OCR Implementation using Tesseract and PyMuPDF

James Sy
July 27, 2022
'''
try:
    from PIL import Image
    import pytesseract
    from cv2 import imread, IMREAD_GRAYSCALE
    from glob import glob
    from fitz import open, Matrix
    from os import remove
    from re import search, IGNORECASE

except ImportError as i:
    print(i.msg)

class Reader:
    '''
    Converts PDFs into sequenced PNGs and uses Tesseract OCR to extract text.
    Pulls out relevant information from this text and stores them in a DocData
    object. Searches the document for a clause matching the user-inputted keyword.

    Class attrs:

    att key1, key2, key3: the user-inputted keywords
    inv: key1, key2, key3 are a string or None

    att files_to_convert: the PDF files to be converted into PNGs and then processed
    inv: files_to_convert is a list of PDF files or None

    att directory: the working directory in which the conversion occurs and the output is given
    inv: directory is an existing directory somewhere on the user's computer

    att ret: the list of DocData objects containing document information
    inv: ret is a list of DocData objects or an empty list

    att files: the list of PNG files in the directory to be OCR'd and parsed
    inv: files is a list of PNG files or None

    att newdoc: the DocData object associated with a PDF
    inv: newdoc is a DocData object or None

    att first_paragraph_trigger: the search word which is the first word of a contract's first paragraph
    inv: first_paragraph_trigger is a non-empty string

    att min_paragraph_spaces: the number of spaces a string must have to be considered a paragraph
    inv: min_paragraph_spaces is an integer
    '''

    def setkey1(self, keyword):
        '''
        Setter for keyword that turns the input into uppercase
        Param: keyword must be a string
        '''
        self.key1 = keyword.upper()

    def setkey2(self, keyword):
        '''
        Setter for keyword that turns the input into uppercase
        Param: keyword must be a string
        '''
        self.key2 = keyword.upper()

    def setkey3(self, keyword):
        '''
        Setter for keyword that turns the input into uppercase
        Param: keyword must be a string
        '''
        self.key3 = keyword.upper()

    def setfiles(self, files):
        '''
        Setter for files that are to be converted
        Param: files must be valid files that are accessible by the user's OS
        '''
        self.files_to_convert = files

    def getdir(self):
        '''
        Returns this Reader's active directory
        '''
        return self.directory

    def setdir(self, directory):
        '''
        Sets this Reader's active directory
        Param: directory must be a valid directory on a path accessible by the user's OS
        '''
        self.directory = directory

    def __init__(self):
        '''
        Initializer
        Param tesser: the path to tesseract.exe on the user's OS
        '''
        try:
            self.key1 = None
            self.key2 = None
            self.key3 = None
            self.files_to_convert = None
            self.directory = None
            self.ret = []
            self.files = None
            self.newdoc = None
            self.first_paragraph_trigger = 'this'
            self.min_paragraph_spaces = 7
            pytesseract.pytesseract.tesseract_cmd = r'tesseract\\tesseract.exe'
        except OSError as e:
            print(e.errno)

    def process(self):
        '''
        Uses Tesseract to extract text from the PNGs in files, combines them into a large string,
        then searches for the keywords inputted by the user. Extracts the first paragraph of the doc,
        and adds everything to a DocData object. Removes the just-searched file after every iteration, 
        and adds each new DocData object to ret.
        '''

        curr = ''

        for file in self.files:

            img = imread(file, IMREAD_GRAYSCALE)
            ship = Image.fromarray(img)
            final = ship.convert('RGB')
            add = pytesseract.image_to_string(final, config='--psm 4')

            curr = curr + add

            remove(file)

        self.newdoc.setpara(self.pull(curr, self.first_paragraph_trigger))
        self.newdoc.set1(self.pull(curr, self.key1))
        self.newdoc.set2(self.pull(curr, self.key2))
        self.newdoc.set3(self.pull(curr, self.key3))

        self.ret.append(self.newdoc)
        self.newdoc = None

    def convert(self):
        '''
        Converts each file of files_to_convert from a PDF to a sequence of PNGs
        by page. Zooms in to achieve better resolution. Initializes a DOcData object
        and stores the file path in it. Saves the PNG images to the specified directory.
        '''
        zoom_x = 2
        zoom_y = 2
        mat = Matrix(zoom_x, zoom_y)

        for filename in self.files_to_convert:

            self.newdoc = DocData(str(filename))

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
        return (self.files_to_convert) != None and \
            (self.directory) != None and \
                (self.key1) != '' and \
                    (self.key1) != None and \
                        (self.key2) != '' and \
                            (self.key2) != None and \
                                (self.key3) != '' and \
                                    (self.key3) != None
            

    def pull(self, document, keyword):
        '''
        Searches for the keyword clause, returns it if found (returns empty string otherwise). 
        Extends to the next chunk if the clause is not complete, and does the same if there
        exists an empty line between the title and the clause.
        Param: document is a string
        Param: keyword is a string
        '''
        try:
            interest = search(keyword, document, IGNORECASE).start()
            slice = document[interest:]
            parabreak = slice.find('\n\n')
            if parabreak != -1:
                chunk = slice[:parabreak]
                if chunk.count(' ') < self.min_paragraph_spaces:
                    nextone = slice[parabreak+1].find('\n\n')
                    chunk = slice[:nextone]
                if chunk[-1] == '.' or chunk[-1] == '\"' or chunk[-1] == chr(8221):
                    chunquito = chunk.strip()
                    return(''.join(chunquito))
                else:
                    keepgoing = slice[parabreak+1:]
                    parabreak2 = keepgoing.find('\n\n')
                    chunk2 = keepgoing[:parabreak2]
                    chunquito = chunk2.strip()
                    return(''.join(chunquito))
        except:
            return ''

    def clear(self):
        '''
        Clear method to restore to defaults
        '''
        self.ret = []
        self.key1 = None
        self.key2 = None
        self.key3 = None
        self.files_to_convert = None
        self.directory = None

class DocData:
    '''
    Class to store data pulled from a contract
    '''

    def getfile(self):
        return self.filename

    def para(self):
        return self.firstparagraph

    def setpara(self, input):
        self.firstparagraph = input

    def input_1(self):
        return self.input1

    def set1(self, input):
        self.input1 = input

    def input_2(self):
        return self.input2

    def set2(self, input):
        self.input2 = input

    def input_3(self):
        return self.input3

    def set3(self, input):
        self.input3 = input

    def __init__(self, file):
        self.filename = file
        self.firstparagraph = None
        self.input1 = None
        self.input2 = None
        self.input3 = None