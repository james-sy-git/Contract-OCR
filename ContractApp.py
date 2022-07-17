'''
GUI to run OCR and Excel-loading functionality
'''
try:
    from Loader import Loader
    from kivy.app import App
    from kivy.uix.boxlayout import BoxLayout
    from kivy.uix.gridlayout import GridLayout
    from kivy.properties import ObjectProperty

except ImportError as i:
    print(i.msg)

class Root(BoxLayout):
    pass
class EverythingElse(GridLayout):

    loader = Loader()
    reader = loader.getreader()
    keyword = None
    xlfilename = None
    xlsheetname = None

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

    def run(self):
        if self.reader.isready():
            print('running!')
            self.loader.excelinit()
            self.loader.excelload()
            self.loader.savewb()
        else:
            print('not ready!')

    def use_xl_setter(self):
        if self.xlfilename != None:
            self.loader.setexcelname(self.xlfilename)
        elif self.xlfilename == None:
            self.loader.setexcelname('newsheet') # default

    def update_filename(self):
        self.xlfilename = self.ids.xlfilename.text

    def use_title_setter(self):
        if self.xlsheetname != None:
            self.loader.setwstitle(self.xlsheetname)
        elif self.xlsheetame == None:
            self.loader.setwstitle('Sheet1') # default

    def update_sheetname(self):
        self.xlsheetname = self.ids.xlsheetname.text
class ContractApp(App):
    
    def build(self):
        root = Root()
        return root

ContractApp().run()