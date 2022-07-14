'''
GUI to run OCR and Excel-loading functionality
'''
try:
    from Loader import Loader
    from kivy.app import App
    from kivy.uix.button import Button
    from kivy.uix.boxlayout import BoxLayout
    from kivy.uix.gridlayout import GridLayout
    from kivy.uix.label import Label

except ImportError as i:
    print(i.msg)

class ContractApp(App):

    def __init__(self):
        App.__init__(self)
        self.root = BoxLayout(orientation='vertical')
        self.header = Label(text = 'CONTRACT PARSER', height=10)
        self.midstack = GridLayout(cols=2, size_hint_y=2)
        self.bottomstack = GridLayout(cols=2, size_hint_y=2)
        self.buttons = BoxLayout(orientation='vertical')
        self.loader = Loader()

    def choosefile(self, press_instance):
        self.loader.getreader().ask_conv()

    def choosedir(self, press_instance):
        self.loader.getreader().ask_dir()

    def loadmidstack(self):
        keylabel = Label(text = 'Keyword')

        selectlabel = Label(text = 'Selected Files')

        self.midstack.add_widget(keylabel)
        self.midstack.add_widget(selectlabel)

    def loadbottomstack(self):
        self.loadbuttons()
        filelabel = Label(text = 'Your Files')

        self.bottomstack.add_widget(self.buttons)
        self.bottomstack.add_widget(filelabel)

    def loadbuttons(self):
        b1 = Button(text='Choose Files')

        # b1.bind(on_press = self.choosefile)

        b2 = Button(text='Choose Directory')

        # b2.bind(on_press = self.choosedir)

        b3 = Button(text = 'File to Write')
        b4 = Button(text = 'Spreadsheet Name')

        self.buttons.add_widget(b1)
        self.buttons.add_widget(b2)
        self.buttons.add_widget(b3)
        self.buttons.add_widget(b4)

    def build(self):

        self.root.add_widget(self.header)
        self.loadmidstack()
        self.loadbottomstack()

        self.root.add_widget(self.midstack)
        self.root.add_widget(self.bottomstack)

        return self.root

ContractApp().run()