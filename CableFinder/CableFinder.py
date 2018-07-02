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
        root.geometry("500x200")
        tk.Label(self, text="ABC Cable Search Thing",padx=10,pady=10,font=("Helvetica", 16)).pack(side="top")
        tk.Label(self, text="Load a new excel spreadsheet to 'Pickel', this is a slow process and will take up 5mins",font=("Helvetica", 8)).pack(side="top")
        self.readSpreadSheet = tk.Button(self,font=("Helvetica", 16))
        self.readSpreadSheet["text"] = "Load a excel spreadsheet"
        self.readSpreadSheet["command"] = self.loadspreadsheet
        self.readSpreadSheet.pack()
        tk.Label(self, text="Load a Pickel Data Stream, very quick",font=("Helvetica", 8)).pack(side="top")
        self.readPickel =tk.Button(self,font=("Helvetica", 16))
        self.readPickel["text"] = "Load a Pickel"
        self.readPickel["command"] = self.pickeled
        self.readPickel.pack()
        self.quit = tk.Button(self, text="QUIT", fg="red",command=root.destroy,font=("Helvetica", 16))
        self.quit.pack(side="bottom")

    def loadspreadsheet(self):
        root.withdraw()
        loadspreadsheet = tk.Toplevel(self)
        data_state = tk.Label(loadspreadsheet, text="Loading Cable Data",padx=10,pady=10,font=("Helvetica", 16))
        data_state.pack()
        file_path = askopenfilename()
        while True:
            try:
                xl = pd.read_excel(file_path,sheet_name=None,)
                break
            except:
                data_state.configure(text = "Sorry that did not work please load a vaild cable spread sheet")
                data_state.update()
        for x in xl:
            xl[x].replace('nan',' ')
        data_state.configure(text = "Data Loaded")                
        tk.Label(loadspreadsheet, text="Please Choose Where to Save the Pickel",padx=10,pady=10,font=("Helvetica", 16)).pack()        
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
        pickle_state = tk.Label(pickeled, text="Please Pick The pickle",padx=10,pady=10,font=("Helvetica", 16))
        pickle_state.pack()
        file_path = askopenfilename()
        ws = pickle.load(open(file_path,'rb'))
        pickle_state.configure(text = "Data Loaded")
        pickle_state.update()
        self.findcable(ws)
        pickeled.destroy()

    def findcable(self,ws):
        findcable = tk.Toplevel(self)
        pickle_state = tk.Label(findcable, text="Please Choose the type of Cable",padx=10,pady=10,font=("Helvetica", 16))
        pickle_state.pack()
        sheets=tk.StringVar()
        for x in ws:
            tk.Radiobutton(findcable, text=x,variable=sheets,value=x,indicatoron=0,width=25).pack()
        tk.Label(findcable, text="Enter Cable Number",padx=10,pady=10,font=("Helvetica", 16)).pack()
        e = Entry(findcable,font=("Helvetica", 16),justify=CENTER)
        e.pack()
        tk.Button(findcable, text="Find Cable",command=lambda: self.results(ws,sheets.get(),e.get()),font=("Helvetica", 16)).pack()       
        tk.Button(findcable, text="QUIT", fg="red",command=root.destroy,font=("Helvetica", 16)).pack()



    def results(self,ws,sheets,cable):
        results = tk.Toplevel(self)
        temp  = ws[sheets]
        temp2 = temp.loc[temp['Cable No'] == int(cable)]
        temp3 = temp2.iat[0,1]
        
        row1 = ['Cable Group',"Cable Number","Cable Type","System Name","Signal" ]
        row2 = [str(sheets),str(cable),str(temp2.iat[0,1]),str(temp2.iat[0,5]),str(temp2.iat[0,10])]
        row3 = ["Source Location","Source Equipment","Source Drawing","Source Facility","Source Connection" ]
        row4 = [str(temp2.iat[0,3]),str(temp2.iat[0,4]),str(temp2.iat[0,9]),str(temp2.iat[0,32]),str(temp2.iat[0,6])]
        row5 = ["Dest Location","Dest Equipment","Dest Drawing","Dest Facility","Dest Connection" ]
        row6 = [str(temp2.iat[0,10]),str(temp2.iat[0,11]),str(temp2.iat[0,16]),str(temp2.iat[0,33]),str(temp2.iat[0,13])]
        
        master = [row1,row2,row3,row4,row5,row6]
        i = 0
        j = 0
        for a in master:
            for b in a:
                if j % 2 == 0:
                    tk.Label(results, text=b,font="-weight bold").grid(row=j,column=i,)
                else:
                    tk.Label(results, text=b).grid(row=j,column=i,)
               
                i = i +1
            i = 0
            j = j+1

       


root = tk.Tk()
app = Application(master=root)
app.mainloop()
