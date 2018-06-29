import os
import tkinter as tk
from tkinter.filedialog import *
import pandas as pd
import pickle

file_path = ""

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        self.create_widgets()

    def create_widgets(self):    
        self.readSpreadSheet = tk.Button(self)
        self.readSpreadSheet["text"] = "Load Cable Data Spreadsheet"
        self.readSpreadSheet["command"] = self.loadspreadsheet
        self.readSpreadSheet.pack(side="top")

        self.readPickel =tk.Button(self)
        self.readPickel["text"] = "Load Pickel"
        self.readPickel["command"] = self.pickeled
        self.readPickel.pack()
        self.quit = tk.Button(self, text="QUIT", fg="red",command=root.destroy)
        self.quit.pack(side="bottom")

    def loadspreadsheet(self):
        root.withdraw()
        loadspreadsheet = tk.Toplevel(self)
        data_state = tk.Label(loadspreadsheet, text="Loading Cable Data")
        data_state.pack()
        file_path = askopenfilename()
        xl = pd.read_excel(file_path,sheet_name=None)                
        data_state.configure(text = "Data Loaded")                
        tk.Label(loadspreadsheet, text="Please Choose Where to Save the Pickel").pack()        
        data_state.update()
        file_path = askopenfilename()
        f = open(file_path,'wb')
        pickle.dump(xl,f)
        f.close()
        root.deiconify()
        loadspreadsheet.destroy()
        
    def pickeled(self):
        root.withdraw()
        pickeled = tk.Toplevel(self)
        pickle_state = tk.Label(pickeled, text="Please Pick The pickle")
        pickle_state.pack()
        file_path = askopenfilename()
        ws = pickle.load(open(file_path,'rb'))
        pickle_state.configure(text = "Data Loaded")
        pickle_state.update()
        self.findcable(ws)
        pickeled.destroy()

    def findcable(self,ws):
        findcable = tk.Toplevel(self)
        pickle_state = tk.Label(findcable, text="Please Choose the type of Cable")
        pickle_state.pack()
        sheets=tk.StringVar()
        for x in ws:
            y = str(x)
            tk.Radiobutton(findcable, text=x,variable=sheets,value=x).pack() 
        e = Entry(findcable)
        e.pack()
        tk.Button(findcable, text="Find Cable",command=lambda: self.results(ws,sheets.get(),e.get())).pack()       
        tk.Button(self, text="QUIT", fg="red",command=root.destroy).pack()



    def results(self,ws,sheets,cable):
        results = tk.Toplevel(self)
        tk.Label(results, text="Cable Type ").grid(row=0, column=0)
        tk.Label(results, text=sheets).grid(row=0, column=1)

        tk.Label(results, text="Cable Number ").grid(row=1, column=0)
        tk.Label(results, text=cable).grid(row=1, column=1)

        tk.Label(results, text="Cable Type ").grid(row=2, column=0)
        temp =ws[sheets]
        temp2 = temp.loc[temp['Cable No'] == int(cable)]
        temp3 = temp2.iat[0,1]
        tk.Label(results, text=temp3).grid(row=2, column=1)

root = tk.Tk()
app = Application(master=root)
app.mainloop()
