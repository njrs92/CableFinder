import os
import tkinter as tk
from tkinter.filedialog import *
import pandas as pd
import pickle
import time


file_path = ""

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        root.geometry("650x250")
        tk.Label(self, text="ABC Cable Search Thing",padx=10,pady=10,font=("Helvetica", 16)).pack(side="top")
        tk.Label(self, text="Load a new excel spreadsheet to 'Pickel', this is a slow process and will take up to 5mins",font=("Helvetica", 12)).pack(side="top")
        self.readSpreadSheet = tk.Button(self,font=("Helvetica", 16))
        self.readSpreadSheet["text"] = "Load a excel spreadsheet"
        self.readSpreadSheet["command"] = self.loadspreadsheet
        self.readSpreadSheet.pack()
        tk.Label(self, text="Load a Pickel Data Stream, this operration is very quick",font=("Helvetica", 12)).pack(side="top")
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
        numOfTrys = 4
        for x in range(0,numOfTrys):
            try:              
                file_path = askopenfilename()
                xl = pd.read_excel(file_path,sheet_name=None,)
                break
            except:
                data_state.configure(text = "Sorry that did not work please load a vaild cable spread sheet")
                data_state.update()
                if x == (numOfTrys-1):
                    data_state.configure(text = "Sorry that failed " + str(x) + " times, closing")
                    data_state.update()
                    time.sleep(5)
                    root.destroy()                   
                    quit(0)
        
        data_state.configure(text = "Excel Data Loaded")                
        data_state.update()
        try:
            temp  = xl['VIDEO']
            temp2 = temp.loc[temp['Cable No'] == 1001]
            temp2.iat[0,1]
            flag = True
        except:
            tk.Label(loadspreadsheet, text="Sorry that appares to be an invalid cable spread sheet",padx=10,pady=10,font=("Helvetica", 16)).pack()
            flag = False
        tk.Button(loadspreadsheet, text="Return", fg="red",command=lambda:[loadspreadsheet.destroy(), root.deiconify()],font=("Helvetica", 16)).pack(side="bottom")
        if flag:
            numOfTrys = 4
            pickel_state = tk.Label(loadspreadsheet, text="Please Choose Where to Save the Pickel",padx=10,pady=10,font=("Helvetica", 16)) 
            pickel_state.pack()
            for x in range(0,numOfTrys):         
                try:
                    file_path = askopenfilename()
                    f = open(file_path,'wb')
                    pickle.dump(xl,f)
                    f.close()
                    flag =False
                    pickel_state.configure(text="Pickel Saved")
                    pickel_state.update()
                    break
                except:
                    pickel_state.configure(text="Sorry unable to save Pickel, try agian")
                    pickel_state.update()
                    if x == (numOfTrys-1):
                        pickel_state.configure(text = "Sorry Pickel failed " + str(x) + " times, unable to proceed" )
                        pickel_state.update()
                       
    def pickeled(self):
        root.withdraw()
        pickeled = tk.Toplevel(self)
        pickle_state = tk.Label(pickeled, text="Please Pick The pickle",padx=10,pady=10,font=("Helvetica", 16))
        pickle_state.pack()
        numOfTrys = 4
        for x in range(0,numOfTrys): 
            try:
                file_path = askopenfilename()
                ws = pickle.load(open(file_path,'rb'))
                pickle_state.configure(text = "Data Loaded")
                pickle_state.update()
                flag = True
                break
            except:
                pickle_state.configure(text = "Sorry Loading the Pickel failed please try again" )
                pickle_state.update()
                if x == (numOfTrys-1):
                        pickle_state.configure(text = "Sorry Loading the Pickel failed " + str(x) + " times, unable to proceed" )
                        pickle_state.update()
                        tk.Button(pickeled, text="Return", fg="red",command=lambda:[pickeled.destroy(), root.deiconify()],font=("Helvetica", 16)).pack(side="bottom")
                        flag = False
        if flag:
            try:
                temp  = ws['VIDEO']
                temp2 = temp.loc[temp['Cable No'] == 1001]
                temp2.iat[0,1]
                pickle_state.configure(text = "Pickle loaded successfully " )
                pickle_state.update()
                tk.Button(pickeled, text="Continue", fg="red",command=lambda:[self.findcable(ws),pickeled.destroy()],font=("Helvetica", 16)).pack(side="bottom")
            except:
                pickle_state.configure(text = "Sorry that pickel appears to invaild" )
                tk.Button(pickeled, text="Return", fg="red",command=lambda:[pickeled.destroy(), root.deiconify()],font=("Helvetica", 16)).pack(side="bottom")


    def findcable(self,ws):
        findcable = tk.Toplevel(self)
        pickle_state = tk.Label(findcable, text="Please Choose the type of Cable",padx=10,pady=10,font=("Helvetica", 16))
        pickle_state.pack()
        sheets=tk.StringVar()
        sheets.set('VIDEO')
        for x in ws:
            tk.Radiobutton(findcable, text=x,variable=sheets,value=x,indicatoron=0,width=25).pack()
        tk.Label(findcable, text="Enter Cable Number",padx=10,pady=10,font=("Helvetica", 16)).pack()
        e = Entry(findcable,font=("Helvetica", 16),justify=CENTER)
        e.pack()
        findcable.tkraise(self)
        e.focus()
        tk.Button(findcable, text="Find Cable",command=lambda: self.results(ws,sheets.get(),e.get()),font=("Helvetica", 16)).pack()       
        tk.Button(findcable, text="QUIT", fg="red",command=root.destroy,font=("Helvetica", 16)).pack()



    def results(self,ws,sheets,cable):
        results = tk.Toplevel(self)
        flag = True
        try: 
            temp  = ws[sheets]
        except:
            tk.Label(results, text="Please select a cable type before clicking find cable" ,padx=10,pady=10,font=("Helvetica", 16)).pack()
            flag = False
            tk.Button(results, text="QUIT", fg="red",command=results.destroy,font=("Helvetica", 16)).pack()  
        
        if flag:            
            try:
                temp2 = temp.loc[temp['Cable No'] == int(cable)]
                temp2.iat[0,1]
            except:
                tk.Label(results, text="Sorry cable " + str(cable) + " in " + str(sheets) + " was not found" ,padx=10,pady=10,font=("Helvetica", 16)).pack()
                flag = False
                tk.Button(results, text="QUIT", fg="red",command=results.destroy,font=("Helvetica", 16)).pack()      
        
        
        if flag:
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
            tk.Button(results, text="QUIT", fg="red",command=results.destroy,font=("Helvetica", 16)).grid(row=7,column=2)
       


root = tk.Tk()
app = Application(master=root)
app.mainloop()
