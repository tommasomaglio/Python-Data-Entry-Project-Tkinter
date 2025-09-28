import tkinter as tk
from tkinter import ttk
import openpyxl 

page = tk.Tk()

def insert_row():
    type = jeweltype.get()
    material = jewelmaterial.get()
    gender = jewelgender.get()
    stockcheck = "In Stock" if var.get() else "Sold Out"

    path = "C:\\Users\\tomma\\Desktop\\Repos\\Python-Data-Entry-Project-Tkinter\\jewel-project.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_val = [type, material, gender, stockcheck]
    sheet.append(row_val)
    workbook.save(path)

    excelview.insert('', tk.END, values=row_val)

    jeweltype.delete(0,"end")
    jewelmaterial.delete(0,"end")
    jewelgender.delete(0,"end")
    jeweltype.insert(0, "Jewel Type")
    jewelmaterial.insert(0, "Jewel Material")
    jewelgender.insert(0,  "Gender")
    insert_button.state(["!selected"])

def load_data():
    path = "C:\\Users\\tomma\\Desktop\\Repos\\Python-Data-Entry-Project-Tkinter\\jewel-project.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    
    rows = [r for r in sheet.values if any(r)]
    if len(rows) < 2:  
        return
    
    headers = [h for h in rows[0] if h is not None]
    excelview["columns"] = headers
    
    for col in headers:
        excelview.heading(col, text=col, anchor="w")
        excelview.column(col, width=100, anchor="w")

    
    for row in rows[1:]:
        clean_row = [cell if cell is not None else "" for cell in row[:len(cols)]]
        excelview.insert("", "end", values=clean_row)

        
        


def switch_toggle():
    if style_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

style = ttk.Style(page)
page.tk.call("source", "forest-light.tcl")
page.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

frame = ttk.Frame(page)
frame.pack()

data_insert_frame = ttk.LabelFrame(frame, text="Insert here:")
data_insert_frame.grid(row=0, column=0, padx=20, pady=10)

types_of_jewels = ['Earring', 'Watch', 'Bracelet', 'Necklace', 'Other']
materials_of_jewels = ['Gold', 'Silver', 'Bronze', 'Metal', 'Titanium', 'Other']
genders_for_jewels = ['Male', 'Female', 'Unisex']

jeweltype = ttk.Combobox(data_insert_frame, values=types_of_jewels)
jeweltype.insert(0, 'Jewel Type')
jeweltype.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

jewelmaterial = ttk.Combobox(data_insert_frame, values=materials_of_jewels)
jewelmaterial.insert(0, 'Jewel Material')
jewelmaterial.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

jewelgender = ttk.Combobox(data_insert_frame, values=genders_for_jewels)
jewelgender.insert(0, 'Gender')
jewelgender.grid(row=2, column=0, sticky="ew", padx=5, pady=5)

var = tk.BooleanVar()
in_stock_check = ttk.Checkbutton(data_insert_frame, text="In Stock", variable=var)
in_stock_check.grid(row=3, column=0, sticky="nsew", padx=5, pady=5)

insert_button = ttk.Button(data_insert_frame, text='Insert', command=insert_row)
insert_button.grid(row=4, column=0, sticky="nsew", padx=5, pady=5)  

separator = ttk.Separator(data_insert_frame)
separator.grid(row=5, column=0, sticky="ew", padx=10, pady=10)

style_switch = ttk.Checkbutton(data_insert_frame, style="Switch", command=switch_toggle)
style_switch.grid(row=6, column=0, sticky="nsew", padx=10, pady=(0, 10))

excelframe = ttk.Frame(frame)
excelframe.grid(row=0, column=1, pady=10)

scrollbar = ttk.Scrollbar(excelframe)   
scrollbar.pack(side="right", fill="y")

cols = ("Type", "Material", "Gender", "Status")
excelview = ttk.Treeview(excelframe, show="headings", columns=cols, height=15,
                         yscrollcommand=scrollbar.set)
excelview.pack()
excelview.column("Type", width=100)
excelview.column("Material", width=100)
excelview.column("Gender", width=100)
excelview.column("Status", width=100)
scrollbar.config(command=excelview.yview)

load_data()

page.mainloop()