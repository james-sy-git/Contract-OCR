'''
GUI to control Reader and Loader functionality
'''

from plistlib import InvalidFileException
import tkinter as tk
from tkinter import ttk
from Loader import Loader
from tkinter.filedialog import askopenfilename, askopenfilenames, askdirectory

class ContractFrame(ttk.Frame):
    '''
    Main working frame that contains Tkinter widgets
    '''

    def __init__(self, container):
        '''
        Initializer
        
        Param container: the root window for this Tkinter Frame
        Pre: container must be a Tk() object 
        '''
        super().__init__(container)

        padding = {'padx': 4, 'pady': 4}

        self.font = ('Segoe UI Emoji', 10) # typeface data

        self.loader = Loader(askopenfilename()) # associated Loader object
        self.reader = self.loader.getreader() # associated Reader object

        self.attrs = [] # buttons and entries

        self.guikeyword = tk.StringVar() # StringVars to store text entry values
        self.filename = tk.StringVar()
        self.sheetname = tk.StringVar()

        self.keyword_label = ttk.Label(self, text='Keyword')
        self.keyword_label.grid(column=0, row=1, **padding)        

        self.keyword_entry = ttk.Entry(self, textvariable=self.guikeyword, width=40)
        self.keyword_entry.grid(column=1, row=1, columnspan=2, **padding)
        self.attrs.append(self.keyword_entry)

        self.file_button = ttk.Button(self, text='Choose Files', width=60,
        command=(lambda: self.ask_conv()))
        self.file_button.grid(column=0, row=2, columnspan=2, **padding)
        self.attrs.append(self.file_button)

        self.directory_button = ttk.Button(self, text='Choose Directory', width=60,
        command=(lambda: self.ask_dir()))
        self.directory_button.grid(column=0, row=3, columnspan=2, **padding)
        self.attrs.append(self.directory_button)

        self.existing_file_button = ttk.Button(self, text='Choose Existing Spreadsheet', width=60,
        command=(lambda: self.askforspreadsheet()))
        self.existing_file_button.grid(column=0, row=4, columnspan=2, **padding)
        self.attrs.append(self.existing_file_button)

        self.excel_file_label = ttk.Label(self, text='Excel File Name')
        self.excel_file_label.grid(column=0, row=5, **padding)

        self.excel_file_entry = ttk.Entry(self, textvariable=self.filename, width=40)
        self.excel_file_entry.grid(column=1, row=5, columnspan=2, **padding)
        self.attrs.append(self.excel_file_entry)

        self.excel_sheet_label = ttk.Label(self, text='Excel Sheet Name')
        self.excel_sheet_label.grid(column=0, row=6, **padding)

        self.excel_sheet_entry = ttk.Entry(self, textvariable=self.sheetname, width=40)
        self.excel_sheet_entry.grid(column=1, row=6, columnspan=2, **padding)
        self.attrs.append(self.excel_sheet_entry)

        self.run_button = ttk.Button(self, text='Convert and Load', 
        command=(lambda: self.runit()), width=60)
        self.run_button.grid(column=0, row=7, columnspan=2, **padding)
        self.attrs.append(self.run_button)

        self.lstyle = ttk.Style()
        self.lstyle.configure('TLabel', font=self.font, background='#F9EF1D')
        self.lstyle.configure('result.TLabel', background='#FFFFFF')
        self.estyle = ttk.Style()
        self.estyle.configure('TEntry', font=self.font)
        self.bstyle = ttk.Style()
        self.bstyle.configure('TButton', font=self.font, background='#F9EF1D')

        self.pack()

    def ask_conv(self):
        '''
        Uses this Frame's reader's setfiles method with a file dialog
        '''
        self.reader.setfiles(askopenfilenames())

    def ask_dir(self):
        '''
        Uses this Frame's reader's setdir method with a file dialog
        '''
        self.reader.setdir(askdirectory())

    def askforspreadsheet(self):
        '''
        Opens file dialog and sets user input as the existing Excel spreadsheet
        '''
        try:
            input = askopenfilename()
            self.loader.setnew(input)
        except InvalidFileException as e:
            print(e.with_traceback()) 

    def lockout(self):
        '''
        Disables all user-inputted widgets in attrs
        '''
        for att in self.attrs:
            att.config(state=tk.DISABLED)
            self.update()

    def restore(self):
        '''
        Enables all user-inputted widgets in attrs
        '''
        for att in self.attrs:
            att.config(state=tk.NORMAL)
            self.update()

    def runit(self):
        '''
        Checks if the Reader object has its fields filled, and runs the contract parser
        '''
        self.lockout()

        self.reader.setkeyword(self.guikeyword.get())
        self.loader.setexcelname(self.filename.get())
        self.loader.setwstitle(self.sheetname.get())

        if self.reader.isready():
            self.loader.excelinit()
            self.loader.excelload()
            self.loader.savewb()
            self.clear()
            
        self.restore()

    def clear(self):
        self.guikeyword.set('')
        self.filename.set('')
        self.sheetname.set('')

class ContractApp(tk.Tk):
    '''
    Container window for the working frame
    '''

    def __init__(self):
        '''
        Initializer
        '''
        super().__init__()

        self.title('Contract Parser')
        self.geometry('440x230')
        self.resizable(False, False)
        self.configure(background='#F9EF1D')

        self.framestyle = ttk.Style()
        self.framestyle.configure('TFrame', background='#F9EF1D')

if __name__ == '__main__':
    app = ContractApp()
    frame = ContractFrame(app)
    app.mainloop() # runs the app