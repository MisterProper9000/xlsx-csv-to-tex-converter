import tkinter as tk
from tkinter import filedialog
import pandas as pd
import csv
import codecs
import xlrd 

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.filename = ""
        self.pack()
        self.create_widgets()
        self.latex_math_list = ["\\", "^", "_"]
        self.latex_escaping_character = ["#", "&", "$"]

    def create_widgets(self):
        self.file_select = tk.Button(self)
        self.file_select["text"] = "Select file"
        self.file_select["command"] = self.select
        self.file_select.pack(side="top")

        self.treat_as_new = tk.BooleanVar()
        self.check_box_threat_as_new = tk.Checkbutton(self,
                                                      text="Treat empty rows as new table sign",
                                                      variable=self.treat_as_new)
        self.check_box_threat_as_new.pack()

        self.remove_borders = tk.BooleanVar()
        self.check_box_remove_borders = tk.Checkbutton(self, 
                                                       text="Remove table borders",
                                                       variable=self.remove_borders)
        self.check_box_remove_borders.pack()

        self.use_math_list = tk.BooleanVar()
        self.check_box_use_math_list = tk.Checkbutton(self,
                                                      text="Use latex math detection",
                                                      variable=self.use_math_list)
        self.check_box_use_math_list.pack()

        self.bottom_caption = tk.BooleanVar()
        self.check_box_bottom_caption = tk.Checkbutton(self,
                                                       text="Table caption at bottom",
                                                       variable=self.bottom_caption)
        self.check_box_bottom_caption.pack()

        self.status = tk.Text(self,state='disabled', width=40, height=3, fg="red")
        self.status.pack()
        self.status.configure(state='normal')
        self.status.insert(tk.END, "No file")
        self.status.configure(state='disabled')
        self.status.tag_config('choosen', background="white", foreground="green")

        self.quit = tk.Button(self, text="Quit", fg="red", command=self.master.destroy)
        self.quit.pack(side="bottom")

        self.choice_delimiter = tk.StringVar()
        self.choice_delimiter_choices = ("current delimiter: \";\"", "current delimiter: \",\"")
        self.choice_delimiter.set(self.choice_delimiter_choices[0])
        self.choice_delimiter_menu = tk.OptionMenu(self, self.choice_delimiter, *self.choice_delimiter_choices)

        self.choice_quotechar = tk.StringVar()
        self.choice_quotechar_choices = ("current quote char: \"", "current quote char: \'")
        self.choice_quotechar.set(self.choice_quotechar_choices[0])
        self.choice_quotechar_menu = tk.OptionMenu(self, self.choice_quotechar, *self.choice_quotechar_choices)

        self.proceed = tk.Button(self, text = "Proceed", command = self.proceed)

    def is_number(self, s):
        try:
            float(s)
            return True
        except ValueError:
            return False

    def select(self):
        try:
            self.proceed.pack_forget()
        except Exception:
            pass
        try:
            self.choice_delimiter_menu.pack_forget()
        except Exception:
            pass
        try:
            self.choice_quotechar_menu.pack_forget()
        except Exception:
            pass

        self.master.filename =  filedialog.askopenfilename(initialdir = "/",
                                                           title = "Select file",
                                                           filetypes = (("xlsx files","*.xlsx"),("csv files","*.csv")))       
        
        file = (self.master.filename.split("/")[-1]).split(".")
        
        if len(file) > 1:
            self.status.configure(state='normal')
            self.status.delete('1.0', tk.END)
            self.status.insert(tk.END, self.master.filename, 'choosen')
            self.status.tag_config('choosen', background="white", foreground="green")

            if file[1] == "csv":                
                self.choice_delimiter_menu.pack();
                self.choice_quotechar_menu.pack();

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
        wb = xlrd.open_workbook(self.master.filename) 
        sheet = wb.sheet_by_index(0) 
        name = (self.master.filename.split("/")[-1]).split(".")[0]
        f = codecs.open('CONVERTER_RESILT.tex','w', "utf-8")
        table_counter = 1
        f.write("\\begin{table}[H]\n"
                + ("\\caption{" + name + str(table_counter) + "}\n" if self.bottom_caption.get() == False else "")
                + "\\label{tab:my_label" + str(table_counter)
                + "}\n\\begin{center}\n\\vspace{5mm}\n\\begin{tabular}{"
                + ("|" if (self.remove_borders.get() == False) else ""))

        actual_column_num = {}
        actual_column_num[table_counter] = sheet.ncols
        length = 0
        counter = 0
        need_new_table = False
        if(self.treat_as_new.get() == True):
            for i in range(sheet.nrows):
                length = 0
                for j in range(sheet.ncols):
                    length += len(str((sheet.cell_value(i,j))).strip())

                if(length == 0):
                    need_new_table = True
                    continue

                if(need_new_table == True):
                    table_counter += 1
                    actual_column_num[table_counter] = 0
                    need_new_table = False

                counter = sheet.ncols
                
                for j in reversed(range(sheet.ncols)):
                    if(len((str(sheet.cell_value(i,j))).strip()) > 0):
                        break
                    else:
                        counter -= 1
                actual_column_num[table_counter] = max(actual_column_num[table_counter], counter)
                
        length = 0
        need_new_table = False
        row_counter = 0
        table_counter = 1

        f.write((actual_column_num[table_counter] - 1) * "c|")
        f.write("c") if self.remove_borders.get() == True else f.write("c|")
        f.write("}\n\\hline\n") if self.remove_borders.get() == False else f.write("}\n")

        for i in range(sheet.nrows):
            length = 0
            for j in range(sheet.ncols):
                length += len((str(sheet.cell_value(i,j))).strip())
            if(length == 0 and self.treat_as_new.get() == True):
                need_new_table = True
                continue

            if(need_new_table == True):
                f.write("\\end{tabular}\n"
                        + ("\\caption{" + name + str(table_counter - 1) + "}\n" if self.bottom_caption.get() == True else "")
                        + "\\end{center}\n\\end{table}\n\n\\begin{table}[H]\n"
                        + ("\\caption{" + name + str(table_counter) + "}\n" if self.bottom_caption.get() == False else "")
                        + "\\label{tab:my_label" + str(table_counter)
                        + "}\n\\begin{center}\n\\vspace{5mm}\n\\begin{tabular}{"
                        + ("|" if self.remove_borders.get() == False else "")
                        + (actual_column_num[table_counter] - 1) * "c|"
                        + ("c|" if self.remove_borders.get() == False else "c")
                        + "}\n" + ("\\hline\n" if self.remove_borders.get() == False else ""))
                need_new_table = False
                table_counter += 1

            for j in range(actual_column_num[table_counter]):
                entry = str(sheet.cell_value(i,j))

                for char in self.latex_escaping_character:
                    entry = entry.replace(char, "\\" + char)

                if (self.use_math_list.get() == True):
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
                            for math in self.latex_math_list:
                                if (sub_entry.find(math) != -1):
                                    sub_entries[index] = "$" + sub_entries[index] + "$"
                                    break

                    entry = ""
                    for sub_entry in sub_entries:
                        entry += sub_entry + " "
                else:
                    for k in range(2, len(self.latex_math_list)):
                        entry = entry.replace(self.latex_math_list[k], "\\" + self.latex_math_list[k])

                if(self.is_number(sheet.cell_value(i,j)) == True and self.use_math_list.get() == True):
                    f.write(" $ " + entry + " $" + " & " if j != (actual_column_num[table_counter] - 1) else "\\\\")
                else:
                    f.write(entry + " & " if j != (actual_column_num[table_counter] - 1) else "\\\\")
            
            length = 0
            for j in range(sheet.ncols):
                length += len((str(sheet.cell_value(min(i + 1, sheet.nrows - 1),j))).strip())

            if((length == 0 and self.treat_as_new.get() == True) or min(i + 1, sheet.nrows - 1) == sheet.nrows - 1):
                if (self.remove_borders.get() == False):
                    f.write("\n\\hline\n")
                else:
                    f.write("\n")
            else:
                f.write("\n\\hline\n")
        
        f.write("\\end{tabular}\n"
                + ("\\caption{" + name + str(table_counter) + "}\n" if self.bottom_caption.get() == True else "")
                + "\\end{center}\n\\end{table}")
        f.close()
       

    def proceedCSV(self):
        pd.set_option('max_colwidth', 40)
        data = pd.read_csv(self.master.filename,sep='\s+', error_bad_lines=False)

        name = (self.master.filename.split("/")[-1]).split(".")[0]

        f = codecs.open('CONVERTER_RESILT.tex','w', "utf-8")
        table_counter = 1
        f.write("\\begin{table}[H]\n"
                + ("\\caption{" + name + str(table_counter) + "}\n" if self.bottom_caption.get() == False else "")
                + "\\label{tab:my_label" + str(table_counter)
                + "}\n\\begin{center}\n\\vspace{5mm}\n\\begin{tabular}{"
                + ("|" if (self.remove_borders.get() == False) else ""))

        actual_column_num = {}
        actual_column_num[table_counter] = data.shape[1]

        delimiter = self.choice_delimiter.get().split("\"")[1]
        quotechar = self.choice_quotechar.get().split(" ")[-1]
        counter = 0;
        need_new_table = False

        # calculate number of columns
        with open(self.master.filename, "r") as my_input_file:
            for row in csv.reader(my_input_file, delimiter=delimiter, quotechar = quotechar):
                counter = 0;

                # if empty line is a table delimiter then no need to count column count across all rows
                if (len(row) == 0 and self.treat_as_new.get() == True):
                    need_new_table = True
                    continue

                if (need_new_table):
                    table_counter += 1
                    actual_column_num[table_counter] = 0
                    need_new_table = False

                for entry in row:
                    counter = counter + 1
                
                actual_column_num[table_counter] = max(actual_column_num[table_counter], counter)

        table_counter = 1
        for i in range(actual_column_num[table_counter] - 1):
            f.write("c|")
        f.write("c") if self.remove_borders.get() == True else f.write("c|")
        f.write("}\n\\hline\n") if self.remove_borders.get() == False else f.write("}\n")

        need_new_table = False
        row_counter = 0

        with open(self.master.filename, "r") as my_input_file:
            cvs_reader = csv.reader(my_input_file, delimiter=delimiter, quotechar = quotechar)
            for row in cvs_reader:
                row_counter = row_counter + 1
                if (len(row) > 0):
                    if (need_new_table == True):
                        table_counter += 1
                        if (self.remove_borders.get() == False):
                            f.write("\n\\hline\n")
                        else:
                            f.write("\n")
                        f.write("\\end{tabular}\n"
                                + ("\\caption{" + name + str(table_counter - 1) + "}\n" if self.bottom_caption.get() == True else "")
                                + "\\end{center}\n\\end{table}\n\n\\begin{table}[H]\n"
                                + ("\\caption{" + name + str(table_counter) + "}\n" if self.bottom_caption.get() == False else "")
                                + "\\label{tab:my_label" + str(table_counter)
                                + "}\n\\begin{center}\n\\vspace{5mm}\n\\begin{tabular}{"
                                + ("|" if self.remove_borders.get() == False else "")
                                + (actual_column_num[table_counter] - 1) * "c|"
                                + ("c|" if self.remove_borders.get() == False else "c")
                                +"}\n" + ("\\hline\n" if self.remove_borders.get() == False else ""))
                        need_new_table = False
                    else:
                        if(row_counter > 1):
                            f.write("\\hline\n")
                    i = 0
                    for entry in row:
                        entry = entry.strip()
                        i += 1

                        for char in self.latex_escaping_character:
                            entry = entry.replace(char, "\\" + char)

                        if (self.use_math_list.get() == True):
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
                                    for math in self.latex_math_list:
                                        if (sub_entry.find(math) != -1):
                                            sub_entries[index] = "$" + sub_entries[index] + "$"
                                            break

                            entry = ""
                            for sub_entry in sub_entries:
                                entry += sub_entry
                        else:
                            for k in range(2, len(self.latex_math_list)):
                                entry = entry.replace(self.latex_math_list[k], "\\" + self.latex_math_list[k])

                        f.write(entry + (" & " if (i != (actual_column_num[table_counter])) else ""))

                    f.write((actual_column_num[table_counter] - i - 1) * "& ")
                    f.write("\\\\")
                    f.write("\n")
                else:
                    if (self.treat_as_new.get() == True):
                        need_new_table = True
                    else:
                        f.write("\n\\hline\n")
                        f.write((actual_column_num[table_counter] - 1) * "& ")
                        f.write("\\\\\n")

            if (self.remove_borders.get() == False):
                f.write("\n\\hline\n")
            else:
                f.write("\n")
            f.write("\\end{tabular}\n"
                    + ("\\caption{" + name + str(table_counter) + "}\n" if self.bottom_caption.get() == True else "")
                    + "\\end{center}\n\\end{table}")

        f.close()


root = tk.Tk()
root.title("XLSX/CSV to TeX converter")
app = Application(master=root)
app.mainloop()