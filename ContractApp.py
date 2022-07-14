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

class Buttons(BoxLayout):
    pass

class ContractApp(App):
    
    def build(self):
        root = Buttons()
        return root

ContractApp().run()