import tkinter as tk
from tkinter import ttk, filedialog, messagebox 
from tksheet import Sheet
import openpyxl as pxl
import pandas as pd
import math

import app
import config

active_index = None
active_type = None
analyzed = []

set_head_product = []
set_head_sale = []

product_sheet_arr = []
sale_sheet_arr = []
nav_sheet_arr = []
def importProductExcel():
    global active_index
    global active_type
    
    try:
        # Load Excel (not include formula, Only read as actual data)
        file_path = filedialog.askopenfilename(filetypes = [("Excel files", "*.xlsx;*.xls")])
        if not file_path:
            return
        data = pd.read_excel(file_path, engine = "openpyxl", header = None).values.tolist()
        headers = data[0]
        data = data[1:]
        
        new_product_sheet = Sheet(app.right_frame,
                                    data = data,
                                    width = 640,
                                    height = 695)
        
        new_product_sheet.headers(headers)
        
        new_sale_sheet = Sheet(app.right_frame,
                                width = 640,
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

        if (isAnyProductData()):
            app.set_product_button["state"] = "active"
            app.import_sale_button["state"] = "active"
        else:
            app.set_product_button["state"] = "disabled"
            app.import_sale_button["state"] = "disabled"

    except Exception as e:
        print(f"Error importing Excel file: {e}")

def importSaleExcel():
    global active_index
    global active_type

    try:
        # Load Excel (not include formula, Only read as actual data)
        file_path = filedialog.askopenfilename(filetypes = [("Excel files", "*.xlsx;*.xls")])
        if not file_path:
            return
        data = pd.read_excel(file_path,engine = "openpyxl",header = None).values.tolist()
        headers = data[0]
        data = data[1:]

        sale_sheet_arr[active_index].set_sheet_data(data) 
        sale_sheet_arr[active_index].headers(headers) 
        
        product_sheet_arr[active_index].grid_remove()
        sale_sheet_arr[active_index].grid_remove()
        sale_sheet_arr[active_index].grid(row=0, column=0, sticky = "nw")
        active_type = "sale"

        if (isAnySaleData()):
            app.set_sale_button["state"] = "active"
        else:
            app.set_sale_button["state"] = "disabled"

    except Exception as e:
        print(f"Error importing Excel file: {e}")

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

    if (isAnyProductData()):
        app.set_product_button["state"] = "active"
        app.import_sale_button["state"] = "active"
    else:
        app.set_product_button["state"] = "disabled"
        app.import_sale_button["state"] = "disabled"

    if (isAnySaleData()):
        app.set_sale_button["state"] = "active"
    else:
        app.set_sale_button["state"] = "disabled"

    print(active_index)
    
def initiateSetProduct():
    set_product_popup = tk.Toplevel(app.app)

    set_product_popup.title("Product Setting")
    set_product_popup.geometry("260x350")
    set_product_popup.minsize(260, 350)
    set_product_popup.maxsize(260, 350)

    set_product_frame = tk.Frame(set_product_popup, bg = "lightblue", padx = 25, pady = 8)
    set_product_frame.grid(row=0, column=0, sticky="n")

    main_label = tk.Label(set_product_frame, text="Specify Key Columns", font=("Open Sans", 10), fg="black")
    main_label.grid(row = 0, column =0, sticky = "w")

    options_headers = product_sheet_arr[active_index].headers()
    display_index_dict = {}
    for index, header in enumerate(product_sheet_arr[active_index].headers()):
        display_index_dict[header] = index

    product_label = tk.Label(set_product_frame, text="Product ID", font=("Open Sans", 10), fg="black")
    product_label.grid(row = 1, column =0, sticky = "w")

    product_combobox = ttk.Combobox(set_product_frame, values = options_headers, state="readonly" )
    product_combobox.grid(row = 2, column = 0, pady = 5)

    price_label = tk.Label(set_product_frame, text="Price", font=("Open Sans", 10), fg="black")
    price_label.grid(row = 3, column =0, sticky = "w")

    price_combobox = ttk.Combobox(set_product_frame, values = options_headers, state="readonly" )
    price_combobox.grid(row = 4, column = 0, pady = 5)

    ltime_label = tk.Label(set_product_frame, text="Lead Time", font=("Open Sans", 10), fg="black")
    ltime_label.grid(row = 5, column =0, sticky = "w")

    ltime_combobox = ttk.Combobox(set_product_frame, values = options_headers, state="readonly" )
    ltime_combobox.grid(row = 6, column = 0, pady = 5)

    fcost_label = tk.Label(set_product_frame, text="Fixed Cost", font=("Open Sans", 10), fg="black")
    fcost_label.grid(row = 7, column =0, sticky = "w")

    fcost_combobox = ttk.Combobox(set_product_frame, values = options_headers, state="readonly" )
    fcost_combobox.grid(row = 8, column = 0, pady = 5)

    vcost_label = tk.Label(set_product_frame, text="Vary Cost", font=("Open Sans", 10), fg="black")
    vcost_label.grid(row = 9, column =0, sticky = "w")

    vcost_combobox = ttk.Combobox(set_product_frame, values = options_headers, state="readonly" )
    vcost_combobox.grid(row = 10, column = 0, pady = 5)
    
    set_product_confirm_button = tk.Button(set_product_popup, text = "Done", width = 10, height = 1, command = lambda: setProduct(display_index_dict[product_combobox.get()], display_index_dict[price_combobox.get()], display_index_dict[ltime_combobox.get()], display_index_dict[fcost_combobox.get()], display_index_dict[vcost_combobox.get()], set_product_popup))
    set_product_confirm_button.place(x = 160, y = 310)
    

def setProduct(product,price,ltime,fcost,vcost,popup):
    global set_head_product

    if (product != "" and price != "" and ltime != "" and fcost != "" and vcost != ""):
        set_head_product = []
        set_head_product.append(product)
        set_head_product.append(price)
        set_head_product.append(ltime)
        set_head_product.append(fcost)
        set_head_product.append(vcost) 
        if (len(set_head_product) == 5):
            app.set_product_button.config(fg="green")
            popup.destroy()
            messagebox.showinfo("Information", "Product key columns are set")
            if (len(set_head_sale) == 3):
                app.analyze_button["state"] = "active"
    

def initiateSetSale():
    set_sale_popup = tk.Toplevel(app.app)

    set_sale_popup.title("Sale Setting")
    set_sale_popup.geometry("260x350")
    set_sale_popup.minsize(260, 240)
    set_sale_popup.maxsize(260, 240)

    set_sale_frame = tk.Frame(set_sale_popup, bg = "lightblue", padx = 25, pady = 8)
    set_sale_frame.grid(row=0, column=0, sticky="n")

    main_label = tk.Label(set_sale_frame, text="Specify Key Columns", font=("Open Sans", 10), fg="black")
    main_label.grid(row = 0, column =0, sticky = "w")
    
    options_headers = sale_sheet_arr[active_index].headers()
    display_index_dict = {}
    for index, header in enumerate(sale_sheet_arr[active_index].headers()):
        display_index_dict[header] = index

    date_label = tk.Label(set_sale_frame, text="Date", font=("Open Sans", 10), fg="black")
    date_label.grid(row = 1, column =0, sticky = "w")

    date_combobox = ttk.Combobox(set_sale_frame, values = options_headers, state="readonly" )
    date_combobox.grid(row = 2, column = 0, pady = 5)

    product_label = tk.Label(set_sale_frame, text="Product ID", font=("Open Sans", 10), fg="black")
    product_label.grid(row = 3, column =0, sticky = "w")

    product_combobox = ttk.Combobox(set_sale_frame, values = options_headers, state="readonly" )
    product_combobox.grid(row = 4, column = 0, pady = 5)

    Q_label = tk.Label(set_sale_frame, text="Quantity", font=("Open Sans", 10), fg="black")
    Q_label.grid(row = 5, column =0, sticky = "w")

    Q_combobox = ttk.Combobox(set_sale_frame, values = options_headers, state="readonly" )
    Q_combobox.grid(row = 6, column = 0, pady = 5)

    set_sale_confirm_button = tk.Button(set_sale_popup, text = "Done", width = 10, height = 1, command = lambda: setSale(display_index_dict[date_combobox.get()], display_index_dict[product_combobox.get()], display_index_dict[Q_combobox.get()], set_sale_popup))
    set_sale_confirm_button.place(x = 160, y = 200)

def setSale(date,product,Q,popup):
    global set_head_sale
    
    if (date != "" and product != "" and Q != ""):
        set_head_sale = []
        set_head_sale.append(date)
        set_head_sale.append(product)
        set_head_sale.append(Q)
        if (len(set_head_sale) == 3):
            app.set_sale_button.config(fg="green")
            popup.destroy()
            messagebox.showinfo("Information", "Sale key columns are set")
            if (len(set_head_product) == 5):
                app.analyze_button["state"] = "active"

def intiateAnalyze():
    analyze_popup = tk.Toplevel(app.app)

    analyze_popup.title("Analyze")
    analyze_popup.geometry("260x220")
    analyze_popup.minsize(260, 220)
    analyze_popup.maxsize(260, 220)

    analyze_frame = tk.Frame(analyze_popup, bg = "lightblue", padx = 25, pady = 8)
    analyze_frame.grid(row=0, column=0, sticky="n")

    case_label = tk.Label(analyze_frame, text="Case", font=("Open Sans", 10), fg="black")
    case_label.grid(row = 0, column =0, sticky = "w")

    options_case = []
    display_index_dict = {}
    for index, case in enumerate(sale_sheet_arr):
        if (isAnyData(case)):
            display = f"Product/ Sale {sale_sheet_arr.index(case) + 1}"
            options_case.append(display)
            display_index_dict[display] = index
    case_combobox = ttk.Combobox(analyze_frame, values = (options_case), state="readonly" )
    case_combobox.grid(row = 1, column = 0, pady = 5)

    period_label = tk.Label(analyze_frame, text="Cycle Period", font=("Open Sans", 10), fg="black")
    period_label.grid(row = 2, column =0, sticky = "w")

    options_period = ["day", "week", "month", "year"]
    period_combobox = ttk.Combobox(analyze_frame, values = options_period, state="readonly" )
    period_combobox.grid(row = 3, column = 0, pady = 5)

    service_label = tk.Label(analyze_frame, text="Service Level (%)", font=("Open Sans", 10), fg="black")
    service_label.grid(row = 4, column =0, sticky = "w")

    service_entry = tk.Entry(analyze_frame, width = 23)
    service_entry.grid(row = 5, column = 0, pady = 5)

    analyze_confirm_button = tk.Button(analyze_popup, text = "Done", width = 10, height = 1, command = lambda: preProcessSheet(display_index_dict[case_combobox.get()], period_combobox.get(), service_entry.get()))
    analyze_confirm_button.place(x = 160, y = 180)

def preProcessSheet(case, period, service):
    print(case)
    print(period)
    print(service)

def isAnyData(sheet):
    data = sheet.get_sheet_data()
    for row in data:
        for cell in row:
            if cell != '':
                print("Data in Table!")
                return True
    print("None in Table!")
    return False

def isAnyProductData():
    global active_index

    if (active_index != None):
        data = product_sheet_arr[active_index].get_sheet_data()
        for row in data:
            for cell in row:
                if cell != '':
                    #print("Data in Table!")
                    return True
        #print("None in Table!")
    return False

def isAnySaleData():
    global active_index
    
    if (active_index != None):
        data = sale_sheet_arr[active_index].get_sheet_data()
        for row in data:
            for cell in row:
                if cell != '':
                    #print("Data in Table!")
                    return True
        #print("None in Table!")
    return False