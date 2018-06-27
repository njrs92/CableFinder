import os
import tkinter as tk
from tkinter.filedialog import askopenfilename
import pandas as pd

file_path = ""

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.hi_there = tk.Button(self)
        self.hi_there["text"] = "load Spreadsheet"
        self.hi_there["command"] = self.result
        self.hi_there.pack(side="top")

        self.quit = tk.Button(self, text="QUIT", fg="red",command=root.destroy)
        self.quit.pack(side="bottom")

    def result(self):
        root.withdraw()
        result = tk.Toplevel(self)
        result.geometry('500x500')

        data_state = tk.Label(result, text="Loading Cable Data")
        data_state.pack()
        file_path = askopenfilename()
        xl = pd.ExcelFile(file_path)
        data_state.configure(text = "Data Loaded")
        tk.Label(result, text="Please Select Type of Cable").pack()
        data_state.update()
        Sheet_choice = tk.StringVar()
        for x in xl.sheet_names:
            tk.Radiobutton(result, text=x, variable=Sheet_choice, value=x).pack()       
        tk.Label(result, text="Enter Cable Number").pack()
        cable_number = tk.Entry(result).pack()
        tk.Button(result, text="Find",command=lambda: self.extract(xl,Sheet_choice)).pack()        
        self.quit = tk.Button(result, text="QUIT", fg="red",command=root.destroy)
        self.quit.pack(side="bottom")
    
    def extract(self,xl,Sheet_choice):
        #df1 = xl.parse(Sheet_choice)
        extract = tk.Toplevel(self)
        tk.Label(extract, text="Sheet Selcted").pack()
        tk.Label(extract, text=Sheet_choice.get()).pack()       
        

root = tk.Tk()
app = Application(master=root)
app.mainloop()
