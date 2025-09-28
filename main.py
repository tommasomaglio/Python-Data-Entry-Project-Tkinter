import tkinter as tk
from tkinter import ttk

page = tk.Tk()

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

jewel_type = ttk.Combobox(data_insert_frame, values=types_of_jewels)
jewel_type.insert(0, 'Jewel Type')
jewel_type.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

jewel_material = ttk.Combobox(data_insert_frame, values=materials_of_jewels)
jewel_material.insert(0, 'Jewel Material')
jewel_material.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

jewel_gender = ttk.Combobox(data_insert_frame, values=genders_for_jewels)
jewel_gender.insert(0, 'Gender')
jewel_gender.grid(row=2, column=0, sticky="ew", padx=5, pady=5)

var = tk.BooleanVar()
in_stock_check = ttk.Checkbutton(data_insert_frame, text="In Stock", variable=var)
in_stock_check.grid(row=3, column=0, sticky="nsew", padx=5, pady=5)

insert_button = ttk.Button(data_insert_frame, text='Insert')
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

page.mainloop()