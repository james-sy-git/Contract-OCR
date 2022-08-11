'''
GUI to control Reader and Loader functionality

James Sy
July 27, 2022
'''

import tkinter as tk
from tkinter import ttk
from Loader import Loader
from tkinter.filedialog import askopenfilenames, askdirectory

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

        self.loader = Loader() # associated Loader object
        self.reader = self.loader.getreader() # associated Reader object

        self.attrs = [] # buttons and entries

        self.keyword1 = tk.StringVar() # StringVars to store text entry values
        self.keyword2 = tk.StringVar()
        self.keyword3 = tk.StringVar()
        self.filename = tk.StringVar()
        self.sheetname = tk.StringVar()

        self.keyword1_label = ttk.Label(self, text='Keyword 1')
        self.keyword1_label.grid(column=0, row=1, **padding)        

        self.keyword1_entry = ttk.Entry(self, textvariable=self.keyword1, width=40)
        self.keyword1_entry.grid(column=1, row=1, columnspan=2, **padding)
        self.attrs.append(self.keyword1_entry)

        self.keyword2_label = ttk.Label(self, text='Keyword 2')
        self.keyword2_label.grid(column=0, row=2, **padding)     

        self.keyword2_entry = ttk.Entry(self, textvariable=self.keyword2, width=40)
        self.keyword2_entry.grid(column=1, row=2, columnspan=2, **padding)
        self.attrs.append(self.keyword2_entry)

        self.keyword3_label = ttk.Label(self, text='Keyword 3')
        self.keyword3_label.grid(column=0, row=3, **padding)     

        self.keyword3_entry = ttk.Entry(self, textvariable=self.keyword3, width=40)
        self.keyword3_entry.grid(column=1, row=3, columnspan=2, **padding)
        self.attrs.append(self.keyword3_entry)        

        self.file_button = ttk.Button(self, text='Choose Files', width=60,
        command=(lambda: self.ask_conv()))
        self.file_button.grid(column=0, row=4, columnspan=2, **padding)
        self.attrs.append(self.file_button)

        self.directory_button = ttk.Button(self, text='Choose Directory', width=60,
        command=(lambda: self.ask_dir()))
        self.directory_button.grid(column=0, row=5, columnspan=2, **padding)
        self.attrs.append(self.directory_button)

        self.excel_file_label = ttk.Label(self, text='Excel File Name')
        self.excel_file_label.grid(column=0, row=6, **padding)

        self.excel_file_entry = ttk.Entry(self, textvariable=self.filename, width=40)
        self.excel_file_entry.grid(column=1, row=6, columnspan=2, **padding)
        self.attrs.append(self.excel_file_entry)

        self.excel_sheet_label = ttk.Label(self, text='Excel Sheet Name')
        self.excel_sheet_label.grid(column=0, row=7, **padding)

        self.excel_sheet_entry = ttk.Entry(self, textvariable=self.sheetname, width=40)
        self.excel_sheet_entry.grid(column=1, row=7, columnspan=2, **padding)
        self.attrs.append(self.excel_sheet_entry)

        self.run_button = ttk.Button(self, text='Convert and Load', 
        command=(lambda: self.runit()), width=60)
        self.run_button.grid(column=0, row=8, columnspan=2, **padding)
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

    def keywordstolist(self, keyword_input):
        '''
        Splits the space-delimited keywords into a list of keywords,
        replaces the underscores with spaces, and returns the new list.
        param: keyword_input is the string of user-inputted keywords
        inv: keyword_input is a list of strings, possibly empty
        '''
        split_it = keyword_input.split()
        ret = []
        for entry in split_it:
            entry = entry.replace('_', ' ')
            ret.append(entry)
        return ret

    def runit(self):
        '''
        Checks if the Reader object has its fields filled, and runs the contract parser
        '''
        self.lockout()

        self.reader.setkey1(self.keywordstolist(self.keyword1.get()))
        self.reader.setkey2(self.keywordstolist(self.keyword2.get()))
        self.reader.setkey3(self.keywordstolist(self.keyword3.get()))
        self.loader.setexcelname(self.filename.get()) 
        self.loader.setwstitle(self.sheetname.get())

        if self.reader.isready():
            self.loader.excelinit(self.keyword1.get(), self.keyword2.get(), self.keyword3.get())
            self.loader.excelload()
            self.loader.savewb()
            self.clear()
            
        self.restore()

    def clear(self):
        '''
        Clears StringVars so that the parser can be used again
        '''
        self.keyword1.set('')
        self.keyword2.set('')
        self.keyword3.set('')
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
        self.geometry('440x260')
        self.resizable(False, False)
        self.configure(background='#F9EF1D')

        self.framestyle = ttk.Style()
        self.framestyle.configure('TFrame', background='#F9EF1D')

if __name__ == '__main__':
    app = ContractApp()
    frame = ContractFrame(app)
    app.mainloop() # runs the app