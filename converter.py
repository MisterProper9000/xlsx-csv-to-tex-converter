import tkinter as tk
from tkinter import filedialog
import pandas as pd
import csv



class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.hi_there = tk.Button(self)
        self.hi_there["text"] = "Select file"
        self.hi_there["command"] = self.select
        self.hi_there.pack(side="top")

        self.treatAsNew = tk.BooleanVar()
        self.checkBox = tk.Checkbutton(self, text="Treat empty rows as new table sign", variable=self.treatAsNew)
        self.checkBox.pack()


        self.proceed = tk.Button(self, text = "Proceed", command = self.proceed)
        self.proceed.pack()

        self.status = tk.Text(self,state='disabled', width=40, height=3, fg="red")
        self.status.pack()
        self.status.configure(state='normal')
        self.status.insert(tk.END, "No file")
        self.status.configure(state='disabled')
        self.status.tag_config('choosen', background="white", foreground="green")

        self.quit = tk.Button(self, text="Quit", fg="red", command=self.master.destroy)
        self.quit.pack(side="bottom")

    def select(self):
        self.master.filename =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xlsx files","*.xlsx"),("csv files","*.csv")))
        self.status.configure(state='normal')
        self.status.delete('1.0', tk.END)
        self.status.insert(tk.END, self.master.filename, 'choosen')
        self.status.configure(state='disabled')

    def proceed(self):
        file = (self.master.filename.split("/")[-1]).split(".")[1]
        if file == "xlsx":
            self.proceedXLSX()
        else:
            self.proceedCSV()


    def proceedXLSX(self):        
        print(":)")

    def proceedCSV(self):
        pd.set_option('max_colwidth', 40)
        data = pd.read_csv(self.master.filename,sep='\s+', error_bad_lines=False) #this is bad but who cares

        name = (self.master.filename.split("/")[-1]).split(".")[0]

        f = open('CONVERTER_RESILT.tex','w')
        tableCounter = 1
        f.write("\\begin{table}[H]\n\\caption{"+name+str(tableCounter)+"}\n\\label{tab:my_label"+str(tableCounter)+"}\n\\begin{center}\n\\vspace{5mm}\n\\begin{tabular}{|")

        for i in range(data.shape[1]):
            f.write("c|")
        f.write("}\n\\hline\n")

        emptyLinesListener = ""
        actualColumnNum = data.shape[1] #csv.reader calculates number of columns by first row, but it could be uncorrect, so this value will be calculated after string parsing

        with open(self.master.filename, "r") as my_input_file:
            for row in csv.reader(my_input_file):
                
                if ((" ".join(row).replace(" ","") == "" or " ".join(row).replace(" ","") == "\n") and self.treatAsNew.get() == True):
                    emptyLinesListener = "1"
                else:
                    if(emptyLinesListener != "1"):
                        string = (" ".join(row)).rstrip(';').replace(";", "&")
                        actualColumnNum = max(actualColumnNum, string.count("&")+1)
                        if string.count("&")+1 < actualColumnNum:
                            string+=(actualColumnNum-1-string.count("&"))*" &"
                            string+="\n"
                        string+="\\\\\n\\hline\n"
                        f.write(string)
                    else:
                        tableCounter += 1
                        string = "\\end{tabular}\n\\end{center}\n\\end{table}\n\n\\begin{table}[H]\n\\caption{"+name+str(tableCounter)+"}\n\\label{tab:my_label"+str(tableCounter)+"}\n\\begin{center}\n\\vspace{5mm}\n\\begin{tabular}{|"+actualColumnNum*"c|"+"}\n\\hline\n"+(" ".join(row)).rstrip(';').replace(";", "&")
                        if string.count("&")+1 < actualColumnNum:
                            string+=(actualColumnNum-1-string.count("&"))*" &"
                            string+="\n"
                        string+="\\\\\n\\hline\n"
                        f.write(string)
                        emptyLinesListener = ""

                   
            
       

            f.write("\\end{tabular}\n\\end{center}\n\\end{table}")

        f.close() 
        print(self.treatAsNew.get())

root = tk.Tk()
root.title("XLSX/CVS to TeX converter")
app = Application(master=root)
app.mainloop()