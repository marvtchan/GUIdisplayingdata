try:
    # Python 2
    import Tkinter as tk
    import ttk
    from tkFileDialog import askopenfilename
except ImportError:
    # Python 3
    import tkinter as tk
    from tkinter import ttk
    from tkinter.filedialog import askopenfilename

import pandas as pd
import numpy as np
from pandas import ExcelFile
from pandas import ExcelWriter

pd.options.display.max_rows = 999

# --- classes ---

class MyWindow:

    def __init__(self, parent):
        self.parent = parent

        self.filename = None
        self.df = None

        self.text = tk.Text(self.parent)
        self.text.pack(pady=10,padx=10)

        self.button = tk.Button(self.parent, text='LOAD DATA', command=self.load)
        self.button.pack(pady=10,padx=10)

        self.button = tk.Button(self.parent, text='DISPLAY DATA', command=self.display)
        self.button.pack(pady=10,padx=10)

        self.button = tk.Button(self.parent, text='SAVE DATA', command=self.save)
        self.button.pack(pady=10,padx=10)

        self.bill_hours = []
        self.emails = None

    #function for extraction of company domains
    def companies(self):
        self.users = self.df['email'].str.split("@", expand= True)
        self.users.columns = ['user','domain']
        self.emails = self.users.domain.unique()
        return self.emails
        print(self.emails)


    #parse files for billable hours
    def weekly_run(self):
        for self.options in self.emails:
            print(self.options)
            self.results = self.df[self.df['email'].str.contains(self.options)]         
            self.companyhours = float(sum(self.results['Hours']))
            print(self.companyhours)
            self.bill_hours.append(self.companyhours)
            print(self.bill_hours)
    
    #function for load button
    def load(self):

        name = askopenfilename(filetypes=[('Excel', ('*.xls', '*.xlsx')), ('CSV', '*.csv',)])

        if name:
            if name.endswith('.csv'):
                self.df = pd.read_csv(name)
            else:
                self.df = pd.read_excel(name)

            self.filename = name
            self.companies()
            print(self.emails)
            self.emails = self.companies()
            self.weekly_run()
            self.result = pd.DataFrame({'Companies': self.emails, 'Hours' : self.bill_hours})
            self.companiesdf1 = self.result.sort_values('Companies')

    #function for display button
    def display(self):
        # ask for file if not loaded yet
        if self.df is None:
            self.load()

        # display if loaded
        if self.df is not None:
            self.text.insert('end', self.filename + '\n')
            self.text.insert('end', str(self.companiesdf1) + '\n')

    #export to excel
    def export(self):
        self.companiesdf1.to_excel('output.xlsx')

    #function for save button
    def save(self):
        # ask for file if not loaded yet
        if self.df is None:
            self.load()

        # display if loaded
        if self.df is not None:
            self.export()

# --- main ---

if __name__ == '__main__':
    root = tk.Tk()
    top = MyWindow(root)
    root.mainloop()