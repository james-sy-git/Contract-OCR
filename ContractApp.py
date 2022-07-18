'''
GUI to run OCR and Excel-loading functionality
'''
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
    pass
class EverythingElse(GridLayout):

    buttondisable = False

    loader = Loader()
    reader = loader.getreader()
    keyword = ''
    xlfilename = None
    xlsheetname = None

    def running(self):
        return self.buttondisable

    def updatekey(self):
        self.keyword = self.ids.key.text

    def setkey(self):
        self.reader.setkeyword(self.keyword)
        print('keyword set:' + self.keyword)

    def ask_conv(self):
        self.reader.ask_conv()
        print('files set')

    def ask_dir(self):
        self.reader.ask_dir()
        print('directory set')

    def askforspreadsheet(self):
        try:
            input = askopenfilename()
            self.loader.setnew(input)
        except:
            print('something is wrong')

    def run(self):
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
        if self.xlfilename != None:
            self.loader.setexcelname(self.xlfilename)

    def update_filename(self):
        self.xlfilename = self.ids.xlfilename.text

    def use_title_setter(self):
        if self.xlsheetname != None:
            self.loader.setwstitle(self.xlsheetname)

    def update_sheetname(self):
        self.xlsheetname = self.ids.xlsheetname.text

    def clear(self):
        self.buttondisable = False
        self.keyword = ''
        self.ids.key.text = ''
        self.xlfilename = None
        self.ids.xlfilename.text = ''
        self.xlsheetname = None
        self.ids.xlsheetname.text = ''

class ContractApp(App):
    
    def build(self):
        root = Root()
        return root

ContractApp().run()