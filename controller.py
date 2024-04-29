import tkinter as tk
from tkinter import filedialog
from tksheet import Sheet
import openpyxl as pxl
import pandas as pd
import math

import app
import config

active_index = None
active_type = None

product_sheet_arr = []
sale_sheet_arr = []
nav_sheet_arr = []
def importExcel():
    global active_index
    global active_type
    
    try:
        # Load Excel (not include formula, Only read as actual data)
        file_path = filedialog.askopenfilename(filetypes = [("Excel files", "*.xlsx;*.xls")])
        if not file_path:
            return  # User cancelled the file dialog

        new_product_sheet = Sheet(app.right_frame,
                                    data = pd.read_excel(file_path,
                                                       engine = "openpyxl",
                                                       header = None).values.tolist(),
                                    width = 670,
                                    height = 695)
        
        new_sale_sheet = Sheet(app.right_frame,
                                width = 670,
                                height = 695)
                          
        
        new_product_sheet.enable_bindings(config.standard_binding)
        new_product_sheet.grid(row=0, column=0, sticky = "nw")

        new_sale_sheet.enable_bindings(config.standard_binding)

        product_sheet_arr.append(new_product_sheet)
        sale_sheet_arr.append(new_sale_sheet)

        active_index = product_sheet_arr.index(new_product_sheet)
        active_type = "product"

        addNavSheet()
        updateNavSheet()

    except Exception as e:
        print(f"Error importing Excel file: {e}")

def importSaleExcel():
    pass

def addNavSheet():
    index = len(nav_sheet_arr)
    new_nav = tk.Button(app.nav_sheet_frame, text = f"Product/Sale {index+1}",width = 12, height = 1, bg="lightblue", command = lambda: selectSheet(index))
    new_nav.grid(row=1, column = index, sticky = "sw")
    nav_sheet_arr.append(new_nav)

def updateNavSheet():
    for nav in nav_sheet_arr:
        nav.grid_remove()

    index = 0
    for nav in nav_sheet_arr:
        nav.grid(row=1, column = index, sticky = "sw")
        index+=1

def selectSheet(selected_sheet_index):
    global active_index
    global active_type

    previous_active = active_index
    product_sheet_arr[active_index].grid_remove()
    sale_sheet_arr[active_index].grid_remove()
        
    if (previous_active != selected_sheet_index):
        product_sheet_arr[selected_sheet_index].grid(row=0, column=0, sticky = "nw")
        active_type = "product"
    else:
        if (active_type == "product"):
            sale_sheet_arr[selected_sheet_index].grid(row=0, column=0, sticky = "nw")
            active_type = "sale"
        elif (active_type == "sale"):
            product_sheet_arr[selected_sheet_index].grid(row=0, column=0, sticky = "nw")
            active_type = "product"
    
    active_index = selected_sheet_index
    print(active_index)

def isAnyData(sheet):
    data = sheet.get_sheet_data()
    for row in data:
        for cell in row:
            if cell != '':
                print("Data in Table!")
                return True
    print("None in Table!")
    return False