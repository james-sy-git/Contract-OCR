'''
GUI to run OCR and Excel-loading functionality
'''
from plistlib import InvalidFileException


try:
    from Loader import Loader
    from kivy.app import App
    from kivy.uix.boxlayout import BoxLayout
    from kivy.uix.gridlayout import GridLayout
    from tkinter.filedialog import askopenfilename
    from tkinter import Tk

    Tk().wm_withdraw()

except ImportError as i:
    print(i.msg)

class Root(BoxLayout):
    '''
    Root widget for Kivy GUI, inherits BoxLayout
    '''
    pass
class EverythingElse(GridLayout):
    '''
    Main widget for Kivy GUI, inherits GridLayout
    
    Class attrs:

    att loader: this EverythingElse object's Loader
    inv: loader is an object of class Loader

    att reader: the Reader associated with this EverythingElse object
    inv: reader is an object of class Reader

    att keyword: the keyword used to search the OCR'd pdf files
    inv: keyword must be a string (initialized as empty string)

    att xlfilename: the desired filename for the saved Excel spreadsheet
    inv: xlfilename must be a string not containing Windows' forbidden characters

    att xlsheetname: the desired sheet name for the populated sheet in the Excel spreadsheet
    inv: xlsheetname must be a string
    '''

    loader = Loader()
    reader = loader.getreader()
    keyword = ''
    xlfilename = None
    xlsheetname = None

    def updatekey(self):
        '''
        Updates keyword according to TextInput field
        '''
        self.keyword = self.ids.key.text

    def setkey(self):
        '''
        Setter for keyword
        '''
        self.reader.setkeyword(self.keyword)
        print('keyword set:' + self.keyword)

    def ask_conv(self):
        '''
        Calls Reader object's ask_conv() method
        '''
        self.reader.ask_conv()
        print('files set')

    def ask_dir(self):
        '''
        Calls Reader object's ask_dir() method
        '''
        self.reader.ask_dir()
        print('directory set')

    def askforspreadsheet(self):
        '''
        Opens file dialog and sets user input as the existing Excel spreadsheet
        '''
        try:
            input = askopenfilename()
            self.loader.setnew(input)
        except InvalidFileException as e:
            print(e.with_traceback())

    def run(self):
        '''
        Checks if the Reader object has its fields filled, and runs the contract parser
        '''
        if self.reader.isready():
            self.buttondisable = True
            print('running!')
            self.loader.excelinit()
            self.loader.excelload()
            self.loader.savewb()
            self.clear()
        else:
            print('not ready!')

    def use_xl_setter(self):
        '''
        Uses Loader's setexcelname() method
        '''
        if self.xlfilename != None:
            self.loader.setexcelname(self.xlfilename)

    def update_filename(self):
        '''
        Updates xlfilename according to TextInput field
        '''
        self.xlfilename = self.ids.xlfilename.text

    def use_title_setter(self):
        '''
        Uses Loader's setwstitle() method
        '''
        if self.xlsheetname != None:
            self.loader.setwstitle(self.xlsheetname)

    def update_sheetname(self):
        '''
        Updates xlsheetname according to TextInput field
        '''
        self.xlsheetname = self.ids.xlsheetname.text

    def clear(self):
        '''
        Clears all populated fields
        '''
        self.buttondisable = False
        self.keyword = ''
        self.ids.key.text = ''
        self.xlfilename = None
        self.ids.xlfilename.text = ''
        self.xlsheetname = None
        self.ids.xlsheetname.text = ''

class ContractApp(App):
    '''
    Class to house and run GUI, extends App
    '''
    
    def build(self):
        root = Root()
        return root

ContractApp().run()