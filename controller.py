import tkinter as tk
from tkinter import ttk, filedialog, messagebox 
from tksheet import Sheet
import openpyxl as pxl
import pandas as pd
import numpy as np
import statistics
from scipy.stats import norm
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
    global set_head_product
    global set_head_sale
    
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

        set_head_product.append([])
        set_head_sale.append([])
        analyzed.append([])

        addNavSheet()
        updateNavSheet()

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

        if (len(set_head_product[active_index]) == 6) and (len(set_head_sale[active_index]) == 3):
            app.analyze_button["state"] = "active"
        else:
            app.analyze_button["state"] = "disabled"

        if (analyzed[active_index] != []):
            app.result_button["state"] = "active"
        else:
            app.result_button["state"] = "disable"

        print(set_head_product)
        print(set_head_sale)
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
    new_nav = tk.Button(app.nav_sheet_frame, text = f"Product/Sale {index+1}",width = 12, height = 1, command = lambda: selectSheet(index))
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
    global set_head_product
    global set_head_sale
    

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

    if (len(set_head_product[active_index]) == 6):
        app.set_product_button.config(fg="green")
    else:
        app.set_product_button.config(fg="black")


    if (len(set_head_sale[active_index]) == 3):
        app.set_sale_button.config(fg="green")
    else:
        app.set_sale_button.config(fg="black")

    if (len(set_head_product[active_index]) == 6) and (len(set_head_sale[active_index]) == 3):
        app.analyze_button["state"] = "active"
    else:
        app.analyze_button["state"] = "disabled"

    if (analyzed[active_index] != []):
        app.result_button["state"] = "active"
    else:
        app.result_button["state"] = "disable"

    print(active_index)
    #print(set_head_product)
    #print(set_head_sale)
    
def initiateSetProduct():
    set_product_popup = tk.Toplevel(app.app)

    set_product_popup.title("Product Setting")
    set_product_popup.geometry("260x350")
    set_product_popup.minsize(260, 400)
    set_product_popup.maxsize(260, 400)

    set_product_frame = tk.Frame(set_product_popup, padx = 25, pady = 8)
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

    hcost_label = tk.Label(set_product_frame, text="Holding Cost", font=("Open Sans", 10), fg="black")
    hcost_label.grid(row = 11, column =0, sticky = "w")

    hcost_combobox = ttk.Combobox(set_product_frame, values = options_headers, state="readonly" )
    hcost_combobox.grid(row = 12, column = 0, pady = 5)
    
    set_product_confirm_button = tk.Button(set_product_popup, text = "Done", width = 10, height = 1, command = lambda: setProduct(display_index_dict[product_combobox.get()], display_index_dict[price_combobox.get()], display_index_dict[ltime_combobox.get()], display_index_dict[fcost_combobox.get()], display_index_dict[vcost_combobox.get()], display_index_dict[hcost_combobox.get()], set_product_popup))
    set_product_confirm_button.place(x = 160, y = 360)
    

def setProduct(product,price,ltime,fcost,vcost,hcost,popup):
    global set_head_product
    global set_head_sale

    if (product != "" and price != "" and ltime != "" and fcost != "" and vcost != "" and hcost != ""):
        set_head_product[active_index] = []
        set_head_product[active_index].append(product)
        set_head_product[active_index].append(price)
        set_head_product[active_index].append(ltime)
        set_head_product[active_index].append(fcost)
        set_head_product[active_index].append(vcost) 
        set_head_product[active_index].append(hcost)
        if (len(set_head_product[active_index]) == 6):
            app.set_product_button.config(fg="green")
            popup.destroy()
            messagebox.showinfo("Information", "Product key columns are set")
            if (len(set_head_sale[active_index]) == 3):
                app.analyze_button["state"] = "active"
        #print(set_head_product)
        #print(set_head_sale)
    

def initiateSetSale():
    set_sale_popup = tk.Toplevel(app.app)

    set_sale_popup.title("Sale Setting")
    set_sale_popup.geometry("260x350")
    set_sale_popup.minsize(260, 240)
    set_sale_popup.maxsize(260, 240)

    set_sale_frame = tk.Frame(set_sale_popup, padx = 25, pady = 8)
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

    sold_Q_label = tk.Label(set_sale_frame, text="Quantity", font=("Open Sans", 10), fg="black")
    sold_Q_label.grid(row = 5, column =0, sticky = "w")

    sold_Q_combobox = ttk.Combobox(set_sale_frame, values = options_headers, state="readonly" )
    sold_Q_combobox.grid(row = 6, column = 0, pady = 5)

    set_sale_confirm_button = tk.Button(set_sale_popup, text = "Done", width = 10, height = 1, command = lambda: setSale(display_index_dict[date_combobox.get()], display_index_dict[product_combobox.get()], display_index_dict[sold_Q_combobox.get()], set_sale_popup))
    set_sale_confirm_button.place(x = 160, y = 200)

def setSale(date,product,sold_Q,popup):
    global set_head_sale
    global set_head_product
    
    if (date != "" and product != "" and sold_Q != ""):
        set_head_sale[active_index] = []
        set_head_sale[active_index].append(date)
        set_head_sale[active_index].append(product)
        set_head_sale[active_index].append(sold_Q)
        if (len(set_head_sale[active_index]) == 3):
            app.set_sale_button.config(fg="green")
            popup.destroy()
            messagebox.showinfo("Information", "Sale key columns are set")
            if (len(set_head_product[active_index]) == 6):
                app.analyze_button["state"] = "active"
        #print(set_head_product)
        #print(set_head_sale)

def intiateAnalyze():
    analyze_popup = tk.Toplevel(app.app)

    analyze_popup.title("Analyze")
    analyze_popup.geometry("260x270")
    analyze_popup.minsize(260, 270)
    analyze_popup.maxsize(260, 270)

    analyze_frame = tk.Frame(analyze_popup, padx = 25, pady = 8)
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

    review_label = tk.Label(analyze_frame, text="Review Period", font=("Open Sans", 10), fg="black")
    review_label.grid(row = 6, column =0, sticky = "w")

    review_entry = tk.Entry(analyze_frame, width = 23)
    review_entry.grid(row = 7, column = 0, pady = 5)

    analyze_confirm_button = tk.Button(analyze_popup, text = "Done", width = 10, height = 1, command = lambda: analyzeSheet(display_index_dict[case_combobox.get()], period_combobox.get(), service_entry.get(), review_entry.get(), analyze_popup))
    analyze_confirm_button.place(x = 160, y = 230)

def analyzeSheet(case, period, service, review, popup):
    global active_index
    global analyzed

    # Product Setting
    product_id_col_prod = set_head_product[active_index][0]
    price_col = set_head_product[active_index][1]
    ltime_col = set_head_product[active_index][2]
    fcost_col = set_head_product[active_index][3]
    vcost_col = set_head_product[active_index][4]
    hcost_col = set_head_product[active_index][5]
    # Sale setting
    date_col = set_head_sale[active_index][0]
    product_id_col_sale = set_head_sale[active_index][1]
    sold_Q_col = set_head_sale[active_index][2]

    data = sale_sheet_arr[case].get_sheet_data()
    # Sum Items/Day/SKU
    period_dict = {}
    for row in data:
        if row[date_col] not in period_dict:
            period_dict[row[date_col]] = {}
        if row[product_id_col_sale] not in period_dict[row[date_col]]:
            period_dict[row[date_col]][row[product_id_col_sale]] = row[sold_Q_col]
        else:
            period_dict[row[date_col]][row[product_id_col_sale]] += row[sold_Q_col]
    
    period_num  = len(period_dict)
    period_text = f"{period_num} {period}"
    #print(period)

    # Service Level -> Safety Factor
    safety_factor = norm.ppf(int(service)/100, loc=0, scale=1)
    #print(safety_factor)

    # Product info.
    data_product = product_sheet_arr[case].get_sheet_data()
    np_matrix = np.array(data_product)
        # Product ID
    data_product_id = np_matrix.T[product_id_col_prod]
    #print(data_product_id)

        # Price
    price_dict = {}
    data_product_price = np_matrix.T[price_col]
    index = 0
    for product_id in data_product_id:
        price_dict[product_id] = data_product_price[index]
        index +=1
    #print(price_dict)

        # Lead time
    ltime_dict = {}
    data_product_leadtime = np_matrix.T[ltime_col]
    index = 0
    for product_id in data_product_id:
        ltime_dict[product_id] = data_product_leadtime[index]
        index +=1
    #print(ltime_dict)

        # Fixed Cost
    fcost_dict = {}
    data_product_fcost = np_matrix.T[fcost_col]
    index = 0
    for product_id in data_product_id:
        fcost_dict[product_id] = data_product_fcost[index]
        index +=1
    #print(fcost_dict)

        # Vary Cost
    vcost_dict = {}
    data_product_vcost = np_matrix.T[vcost_col]
    index = 0
    for product_id in data_product_id:
        vcost_dict[product_id] = data_product_vcost[index]
        index +=1
    #print(vcost_dict)

        # Holding Cost
    hcost_dict = {}
    data_product_hcost = np_matrix.T[hcost_col]
    index = 0
    for product_id in data_product_id:
        hcost_dict[product_id] = data_product_hcost[index]
        index +=1
    #print(hcost_dict)

    # Calculate AVG Demand
    avg_dict = {}
    for product_id in data_product_id:
        product_cum = 0
        for key in period_dict:
            if product_id in period_dict[key]:
                product_cum += period_dict[key][product_id]
        product_avg = product_cum/ period_num  
        avg_dict[product_id] = product_avg  
    #print(avg_dict) 

    # Calculate AVG Demand during Lead Time
    avgl_dict = {}
    for product_id in data_product_id:
        product_avgl = float(avg_dict[product_id])*float(ltime_dict[product_id])
        avgl_dict[product_id] = product_avgl
    #print(avgl_dict) 

    # Calculate STD Demand
    std_dict = {}
    for product_id in data_product_id:
        product_date_arr = []
        for key in period_dict:
            if product_id in period_dict[key]:
                product_date_arr.append(period_dict[key][product_id])
            else:
                product_date_arr.append(0)
        product_std = statistics.stdev(product_date_arr) 
        std_dict[product_id] = product_std 
    #print(std_dict)

    # Safety Stock
    safety_stock_dict = {}
    for product_id in data_product_id:
        safety_stock = safety_factor * float(std_dict[product_id]) * math.sqrt(float(ltime_dict[product_id])) 
        safety_stock_dict[product_id] = math.ceil(safety_stock)
    #print(safety_stock_dict)

    # (Q,R) Policy
        # Order Quantity
    order_Q_dict = {}
    for product_id in data_product_id:
        order_Q = math.sqrt(((2 * float(fcost_dict[product_id])) * float(avg_dict[product_id])) / float(hcost_dict[product_id])) 
        order_Q_dict[product_id] = math.ceil(order_Q)   
    #print(order_Q_dict)

        # Reorder Level
    reorder_dict = {}
    for product_id in data_product_id:
        reorder = float(avgl_dict[product_id]) + float(safety_stock_dict[product_id])
        reorder_dict[product_id] = math.ceil(reorder)   
    #print(reorder_dict)

        # Average Inventory
    avgi_dict = {}
    for product_id in data_product_id:
        avgi = (float(order_Q_dict[product_id])/2) + float(safety_stock_dict[product_id])
        avgi_dict[product_id] = avgi
    #print(avgi_dict)

    # (s,S) Policy
        # s
    s_dict = {}
    for product_id in data_product_id:
        s = reorder_dict[product_id]
        s_dict[product_id] = s
    #print(s_dict)

        # S
    S_dict = {}
    for product_id in data_product_id:
        S = reorder_dict[product_id] + order_Q_dict[product_id]
        S_dict[product_id] = S
    #print(S_dict) 

    # Base Stock Policy
        # Average demand during an interval of r + L days
    avgl_bs_dict = {}
    for product_id in data_product_id:
        avgl_bs = (float(review) + float(ltime_dict[product_id])) * float(avg_dict[product_id])
        avgl_bs_dict[product_id] = avgl_bs
    #print(avgl_bs_dict)

        # Safety Stock (Base Stock Policy)
    safety_stock_bs_dict = {}
    for product_id in data_product_id:
        safety_stock_bs = safety_factor * float(std_dict[product_id]) * math.sqrt((float(review) + float(ltime_dict[product_id]))) 
        safety_stock_bs_dict[product_id] = math.ceil(safety_stock_bs)
    #print(safety_stock_bs_dict)

        # Base-stock level
    bs_level_dict = {}
    for product_id in data_product_id:
        bs_level = float(avgl_bs_dict[product_id]) + float(safety_stock_bs_dict[product_id])
        bs_level_dict[product_id] = math.ceil(bs_level)   
    #print(bs_level_dict)

        # Average inventory (Base Stock Policy)
    avgi_bs_dict = {}
    for product_id in data_product_id:
        avgi_bs = (float(review)*float(avg_dict[product_id]))/2 + float(safety_stock_bs_dict[product_id])
        avgi_bs_dict[product_id] = avgi_bs 
    #print(avgi_bs_dict)

    analyzed[active_index] = {# Processed Input
                            "data_product_id": data_product_id, # array
                            "period_dict": period_dict, 
                            "period_text": period_text, # string
                            "safety_factor": safety_factor, #float
                            # Processed product
                            "avg_dict": avg_dict, 
                            "avgl_dict": avgl_dict, 
                            "std_dict": std_dict, 
                            "safety_stock_dict": safety_stock_dict, 
                                # (Q,R)
                            "order_Q_dict": order_Q_dict, 
                            "reorder_dict": reorder_dict,
                            "avgi_dict": avgi_dict, 
                                # (s,S)
                            "s_dict": s_dict, 
                            "S_dict": S_dict,
                                # Base-stock
                            "avgl_bs_dict": avgl_bs_dict, 
                            "safety_stock_bs_dict": safety_stock_bs_dict, 
                            "bs_level_dict": bs_level_dict,
                            "avgi_bs_dict": avgi_bs_dict}
    
    if (analyzed[active_index] != []):
        app.result_button["state"] = "active"
    else:
        app.result_button["state"] = "disable"

    popup.destroy()
    messagebox.showinfo("Information", "Result is saved")

def result():
    global analyzed

    result_popup = tk.Toplevel(app.app)

    result_popup.title("Result")
    result_popup.geometry("560x750")
    result_popup.minsize(560, 750)
    result_popup.maxsize(560, 750)

    result_label = tk.Label(result_popup, text=f"Result: Product / Sale {active_index + 1}", font=("Open Sans", 20), fg="black")
    result_label.grid(row = 0, column =0, sticky = "nw")

    result_frame = tk.Frame(result_popup, padx = 25, pady = 8)
    result_frame.grid(row=1, column=0, sticky="w")

    input_process_label = tk.Label(result_frame, text=f"Processed Input", font=("Open Sans", 16), fg="black")
    input_process_label.grid(row = 0, column = 0, sticky = "w")

    process_input_frame = tk.Frame(result_frame, padx = 25, pady = 12)
    process_input_frame.grid(row=1, column=0, sticky="w")

    product_num_label = tk.Label(process_input_frame, text= f"Number of products: {len(analyzed[active_index]['data_product_id'])}", font=("Open Sans", 10), fg="black")
    product_num_label.grid(row = 0, column = 0, sticky = "w", pady = 12)

    period_label = tk.Label(process_input_frame, text= f"Period: {analyzed[active_index]['period_text']}", font=("Open Sans", 10), fg="black")
    period_label.grid(row = 0, column = 1, sticky = "w", padx = (15,0), pady = 12)

    safetyf_label = tk.Label(process_input_frame, text= f"Safety Factor: {round(analyzed[active_index]['safety_factor'], 2)}", font=("Open Sans", 10), fg="black")
    safetyf_label.grid(row = 1, column = 0, sticky = "w", pady = 12)

    product_process_label = tk.Label(result_frame, text=f"Processed Product", font=("Open Sans", 16), fg="black")
    product_process_label.grid(row = 2, column = 0, sticky = "w", pady = 12)

    options_product_process = []
    for product in analyzed[active_index]['data_product_id']:
        options_product_process.append(product)
    product_process_combobox = ttk.Combobox(result_frame, values = options_product_process, state="readonly" )
    product_process_combobox.bind('<<ComboboxSelected>>',lambda event: update_product_labels(process_product_frame, product_process_combobox.get()))

    process_product_frame = tk.Frame(result_frame, padx = 25, pady = 12)

    product_process_combobox.grid(row = 3, column = 0, pady = 10, sticky="w")
    process_product_frame.grid(row=4, column=0, sticky="w")


def update_product_labels(frame, selected):

    for widget in frame.winfo_children():
        widget.grid_remove()
        
    selected_product_option = selected
  
    avg_label = tk.Label(frame, text= f"Average Demand: {round(analyzed[active_index]['avg_dict'][selected_product_option],2)}", font=("Open Sans", 10), fg="black")
    avg_label.grid(row = 0, column = 0, sticky = "w", pady = 10)

    avgl_label = tk.Label(frame, text= f"Average Demand during Lead Time: {round(analyzed[active_index]['avgl_dict'][selected_product_option],2)}", font=("Open Sans", 10), fg="black")
    avgl_label.grid(row = 0, column = 1, sticky = "w", padx = (15,0), pady = 10)

    std_label = tk.Label(frame, text= f"Standard Deviation of Demand: {round(analyzed[active_index]['std_dict'][selected_product_option],2)}", font=("Open Sans", 10), fg="black")
    std_label.grid(row = 1, column = 0, sticky = "w", pady = 10)

    safety_stock_label = tk.Label(frame, text= f"Safety Stock: {analyzed[active_index]['safety_stock_dict'][selected_product_option]}", font=("Open Sans", 10), fg="black")
    safety_stock_label.grid(row = 1, column = 1, sticky = "w", padx = (15,0), pady = 10)

    QR_label = tk.Label(frame, text=f"(Q,R) Policy", font=("Open Sans", 14), fg="black")
    QR_label.grid(row = 2, column = 0, sticky = "w", pady = 10)

    order_Q_label = tk.Label(frame, text= f"Order Quantity (Q): {analyzed[active_index]['order_Q_dict'][selected_product_option]}", font=("Open Sans", 10), fg="black")
    order_Q_label.grid(row = 3, column = 0, sticky = "w", pady = 10)

    reorder_label = tk.Label(frame, text= f"Reorder Level (R): {analyzed[active_index]['reorder_dict'][selected_product_option]}", font=("Open Sans", 10), fg="black")
    reorder_label.grid(row = 3, column = 1, sticky = "w", padx = (15,0), pady = 10)

    avgi_label = tk.Label(frame, text= f"Average Inventory: {round(analyzed[active_index]['avgi_dict'][selected_product_option],2)}", font=("Open Sans", 10), fg="black")
    avgi_label.grid(row = 4, column = 0, sticky = "w", pady = 10)

    sS_label = tk.Label(frame, text=f"(s,S) Policy", font=("Open Sans", 14), fg="black")
    sS_label.grid(row = 5, column = 0, sticky = "w", pady = 10)

    s_label = tk.Label(frame, text= f"s: {analyzed[active_index]['s_dict'][selected_product_option]}", font=("Open Sans", 10), fg="black")
    s_label.grid(row = 6, column = 0, sticky = "w", pady = 10)

    S_label = tk.Label(frame, text= f"S: {analyzed[active_index]['S_dict'][selected_product_option]}", font=("Open Sans", 10), fg="black")
    S_label.grid(row = 6, column = 1, sticky = "w", padx = (15,0), pady = 10)

    bs_label = tk.Label(frame, text=f"Base-stock Policy", font=("Open Sans", 14), fg="black")
    bs_label.grid(row = 7, column = 0, sticky = "w", pady = 10)

    avgl_bs_label = tk.Label(frame, text= f"Average Demand during (r + L) Time: {round(analyzed[active_index]['avgl_bs_dict'][selected_product_option],2)}", font=("Open Sans", 10), fg="black")
    avgl_bs_label.grid(row = 8, column = 0, sticky = "w", pady = 10)

    safety_stock_bs_label = tk.Label(frame, text= f"Safety Stock: {analyzed[active_index]['safety_stock_bs_dict'][selected_product_option]}", font=("Open Sans", 10), fg="black")
    safety_stock_bs_label.grid(row = 8, column = 1, sticky = "w", padx = (15,0), pady = 12)

    bs_level_label = tk.Label(frame, text= f"Base-stock Level: {analyzed[active_index]['bs_level_dict'][selected_product_option]}", font=("Open Sans", 10), fg="black")
    bs_level_label.grid(row = 9, column = 0, sticky = "w", pady = 10)

    avgi_bs_label = tk.Label(frame, text= f"Average Inventory: {round(analyzed[active_index]['avgi_bs_dict'][selected_product_option],2)}", font=("Open Sans", 10), fg="black")
    avgi_bs_label.grid(row = 9, column = 1, sticky = "w", padx = (15,0), pady = 10)


def isAnyData(sheet):
    data = sheet.get_sheet_data()
    for row in data:
        for cell in row:
            if cell != '':
                return True
    return False

def isAnyProductData():
    global active_index

    if (active_index != None):
        data = product_sheet_arr[active_index].get_sheet_data()
        for row in data:
            for cell in row:
                if cell != '':
                    return True
    return False

def isAnySaleData():
    global active_index
    
    if (active_index != None):
        data = sale_sheet_arr[active_index].get_sheet_data()
        for row in data:
            for cell in row:
                if cell != '':
                    return True
    return False