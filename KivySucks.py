'''
GUI to control Reader and Loader functionality
'''

from plistlib import InvalidFileException
import tkinter as tk
from tkinter import ttk
from Loader import Loader
from tkinter.filedialog import askopenfilename

class ContractFrame(ttk.Frame):

    def __init__(self, container):
        super().__init__(container)

        padding = {'padx': 4, 'pady': 4}

        self.font = ('Segoe UI Emoji', 8)

        self.loader = Loader()
        self.reader = self.loader.getreader()

        self.keyword = tk.StringVar()
        self.filename = tk.StringVar()
        self.sheetname = tk.StringVar()

        self.keyword_label = ttk.Label(self, text='Keyword')
        self.keyword_label.grid(column=0, row=1, **padding)        

        self.keyword_entry = ttk.Entry(self, textvariable=self.keyword, width=40)
        self.keyword_entry.grid(column=1, row=1, columnspan=2, **padding)
        self.keyword_entry.focus()

        self.file_button = ttk.Button(self, text='Choose Files', width=60,
        command=(lambda: self.reader.ask_conv()))
        self.file_button.grid(column=0, row=2, columnspan=2, **padding)

        self.directory_button = ttk.Button(self, text='Choose Directory', width=60,
        command=(lambda: self.reader.ask_dir()))
        self.directory_button.grid(column=0, row=3, columnspan=2, **padding)

        self.existing_file_button = ttk.Button(self, text='Choose Existing Spreadsheet', width=60,
        command=(lambda: self.askforspreadsheet()))
        self.existing_file_button.grid(column=0, row=4, columnspan=2, **padding)

        self.excel_file_label = ttk.Label(self, text='Excel File Name')
        self.excel_file_label.grid(column=0, row=5, **padding)

        self.excel_file_entry = ttk.Entry(self, textvariable=self.filename, width=40)
        self.excel_file_entry.grid(column=1, row=5, columnspan=2, **padding)

        self.excel_sheet_label = ttk.Label(self, text='Excel Sheet Name')
        self.excel_sheet_label.grid(column=0, row=6, **padding)

        self.excel_sheet_entry = ttk.Entry(self, textvariable=self.sheetname, width=40)
        self.excel_sheet_entry.grid(column=1, row=6, columnspan=2, **padding)

        self.run_button = ttk.Button(self, text='Convert and Load', 
        command=(lambda: self.runit()), width=60)
        self.run_button.grid(column=0, row=7, columnspan=2, **padding)

        self.quitbutton = ttk.Button(self, text='Quit', command=quit, width=10)
        self.quitbutton.grid(column=0, row=9, columnspan=2, **padding)

        self.pack()   

    def askforspreadsheet(self):
        '''
        Opens file dialog and sets user input as the existing Excel spreadsheet
        '''
        try:
            input = askopenfilename()
            self.loader.setnew(input)
        except InvalidFileException as e:
            print(e.with_traceback()) 

    def runit(self):
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

class ContractApp(tk.Tk):

    def __init__(self):
        super().__init__()

        self.title('Contract Parser')
        self.geometry('400x260')
        self.resizable(False, False)
        self.configure(background='#39cf1b')

        self.framestyle = ttk.Style()
        self.framestyle.configure('TFrame', background='#B8FFFD')

if __name__ == '__main__':
    app = ContractApp()
    frame = ContractFrame(app)
    app.mainloop()