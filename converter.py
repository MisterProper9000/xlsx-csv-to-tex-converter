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
        self.latexMathList = ["\\", "^", "_"]
        self.latexEscapingCharacter = ["#", "&"]
        self.quotechar = '\"'

    def create_widgets(self):
        self.hi_there = tk.Button(self)
        self.hi_there["text"] = "Select file"
        self.hi_there["command"] = self.select
        self.hi_there.pack(side="top")

        self.treatAsNew = tk.BooleanVar()
        self.checkBox = tk.Checkbutton(self, text="Treat empty rows as new table sign", variable=self.treatAsNew)
        self.checkBox.pack()

        self.removeBorders = tk.BooleanVar()
        self.checkBox1 = tk.Checkbutton(self, text="Remove table borders", variable=self.removeBorders)
        self.checkBox1.pack()

        self.useMathList = tk.BooleanVar()
        self.checkBox2 = tk.Checkbutton(self, text="Use latex math detection", variable=self.useMathList)
        self.checkBox2.pack()

        self.bottomCaption = tk.BooleanVar()
        self.checkBox2 = tk.Checkbutton(self, text="Table caption at bottom", variable=self.bottomCaption)
        self.checkBox2.pack()

        self.status = tk.Text(self,state='disabled', width=40, height=3, fg="red")
        self.status.pack()
        self.status.configure(state='normal')
        self.status.insert(tk.END, "No file")
        self.status.configure(state='disabled')
        self.status.tag_config('choosen', background="white", foreground="green")

        self.quit = tk.Button(self, text="Quit", fg="red", command=self.master.destroy)
        self.quit.pack(side="bottom")

    def is_number(self, s):
        try:
            float(s)
            return True
        except ValueError:
            return False

    def select(self):
        self.master.filename =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xlsx files","*.xlsx"),("csv files","*.csv")))       
        print(self.master.filename)
        file = (self.master.filename.split("/")[-1]).split(".")
        if len(file) > 1:
            self.status.configure(state='normal')
            self.status.delete('1.0', tk.END)
            self.status.insert(tk.END, self.master.filename, 'choosen')
            self.status.configure(state='disabled')

            if file[1] == "csv":                
                self.choiceDelimiter = tk.StringVar()
                self.choices = ("current delimiter: \";\"", "current delimiter: \",\"")
                self.choiceDelimiter.set(self.choices[0])
                self.w = tk.OptionMenu(self, self.choiceDelimiter, *self.choices)
                self.w.pack();

            self.proceed = tk.Button(self, text = "Proceed", command = self.proceed)
            self.proceed.pack()
        else:
            self.status.configure(state='normal')
            self.status.delete('1.0', tk.END)
            self.status.insert(tk.END, "No file", 'choosen')
            self.status.configure(state='disabled')
            self.status.tag_config('choosen', background="white", foreground="red")


    def proceed(self):
        file = (self.master.filename.split("/")[-1]).split(".")[1]
        if file == "xlsx":
            self.proceedXLSX()
        else:
            self.proceedCSV()
        self.status.configure(state='normal')
        self.status.delete('1.0', tk.END)
        self.status.insert(tk.END, "File processed. Result saved in CONVERTER_RESILT.tex", 'choosen')
        self.status.configure(state='disabled') 
        


    def proceedXLSX(self):        
        print(":)")

    def proceedCSV(self):
        pd.set_option('max_colwidth', 40)
        data = pd.read_csv(self.master.filename,sep='\s+', error_bad_lines=False) #this is bad but who cares

        name = (self.master.filename.split("/")[-1]).split(".")[0]

        f = open('CONVERTER_RESILT.tex','w')
        tableCounter = 1
        f.write("\\begin{table}[H]\n"+("\\caption{"+name+str(tableCounter)+"}\n" if self.bottomCaption.get() == False else "")+"\\label{tab:my_label"+str(tableCounter)+"}\n\\begin{center}\n\\vspace{5mm}\n\\begin{tabular}{" + ("|" if (self.removeBorders.get() == False) else ""))
        
        emptyLinesListener = ""
        actualColumnNum = {}
        actualColumnNum[tableCounter] = data.shape[1] #csv.reader calculates number of columns by first row, but it could be uncorrect, so this value will be calculated after string parsing
        

        delimiter = self.choiceDelimiter.get().split("\"")[1]
        counter = 0;
        needNewTable = False
        with open(self.master.filename, "r") as my_input_file: #calculate number of columns
            for row in csv.reader(my_input_file, delimiter=delimiter, quotechar = self.quotechar):
                #print(row)
                counter = 0;
                if (len(row) == 0 and self.treatAsNew.get() == True):#if empty line is a table delimiter then no need to count column count across all rows
                    needNewTable = True
                    continue

                if (needNewTable):
                    tableCounter += 1
                    actualColumnNum[tableCounter] = 0
                    needNewTable = False

                for entry in row:
                    if len(entry.strip()) > 0:
                        counter = counter + 1
                
                actualColumnNum[tableCounter] = max(actualColumnNum[tableCounter], counter)
                #print(tableCounter)
                #print(actualColumnNum[tableCounter])
                #print("______")

        tableCounter = 1
        for i in range(actualColumnNum[tableCounter]-1):
            f.write("c|")
        f.write("c") if self.removeBorders.get() == True else f.write("c|")
        f.write("}\n\\hline\n") if self.removeBorders.get() == False else f.write("}\n")
        #print(len(actualColumnNum))
        
        
        needNewTable = False
        row_counter = 0
        
        with open(self.master.filename, "r") as my_input_file:
            cvs_reader = csv.reader(my_input_file, delimiter=delimiter, quotechar = self.quotechar)
            for row in cvs_reader:
                #print(row)
                row_counter = row_counter + 1
                if (len(row) > 0):
                    if (needNewTable == True):
                        tableCounter += 1
                        if (self.removeBorders.get() == False):
                            f.write("\n\\hline\n")
                        else:
                            f.write("\n")
                        f.write("\\end{tabular}\n"+("\\caption{"+name+str(tableCounter-1)+"}\n" if self.bottomCaption.get() == True else "")+"\\end{center}\n\\end{table}\n\n\\begin{table}[H]\n"+("\\caption{"+name+str(tableCounter)+"}\n" if self.bottomCaption.get() == False else "")+"\\label{tab:my_label"+str(tableCounter)+"}\n\\begin{center}\n\\vspace{5mm}\n\\begin{tabular}{"+("|" if self.removeBorders.get() == False else "")+(actualColumnNum[tableCounter]-1)*"c|"+("c|" if self.removeBorders.get() == False else "c")+"}\n"+("\\hline\n" if self.removeBorders.get() == False else ""))
                        needNewTable = False
                    else:
                        if(row_counter > 1):
                            f.write("\\hline\n")
                    i = 0
                    for entry in row:

                        entry = entry.strip()
                        if(len(entry) == 0):
                            continue
                        i += 1

                        for char in self.latexEscapingCharacter:
                            entry = entry.replace(char, "\\" + char)

                        if (self.useMathList.get() == True):
                            sub_entries = entry.split(" ")
                            for index, sub_entry in enumerate(sub_entries):
                                if (self.is_number(sub_entry) == True):
                                    sub_entries[index] = "$" + sub_entry
                                    j = 1
                                    while(index + j < len(sub_entries) and self.is_number(sub_entries[j]) == True):
                                        sub_entries[index] += " " + sub_entries[index + j]
                                        sub_entries[index + j] = ""
                                        j = j + 1
                                    sub_entries[index] += "$"
                                    index += j - 1
                                    continue
                                else:
                                    for math in self.latexMathList:
                                        if (sub_entry.find(math) != -1):
                                            sub_entries[index] = "$" + sub_entries[index] + "$"
                                            break

                            entry = ""
                            for sub_entry in sub_entries:
                                entry += sub_entry

                        f.write(entry + (" & " if (i != (actualColumnNum[tableCounter])) else ""))
                    
                    f.write((actualColumnNum[tableCounter]-i-1)*"& ")
                    f.write("\\\\")
                    f.write("\n")
                else:
                    if (self.treatAsNew.get() == True):
                        needNewTable = True
                    else:
                        f.write("\n\\hline\n")
                        f.write((actualColumnNum[tableCounter] - 1) * "& ")
                        f.write("\\\\\n")

            if (self.removeBorders.get() == False):
                f.write("\n\\hline\n")
            else:
                f.write("\n")
            f.write("\\end{tabular}\n"+("\\caption{"+name+str(tableCounter)+"}\n" if self.bottomCaption.get() == True else "")+"\\end{center}\n\\end{table}")

        f.close()

root = tk.Tk()
root.title("XLSX/CSV to TeX converter")
app = Application(master=root)
app.mainloop()