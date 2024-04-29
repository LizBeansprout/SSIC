import tkinter as tk
from tkinter import ttk
from tksheet import Sheet
import config
import controller

# Intiate tkinter entity
app = tk.Tk()

# Setting window size
app.geometry("1080x720")
app.minsize(1080,720)
app.maxsize(1080,720)

# Setting title
app.title("Single-Stage Inventory Control")

# Creating menu bar
menu_bar = tk.Menu(app)
app.config(menu = menu_bar)

# Adding major menu
file_menu = tk.Menu(menu_bar, tearoff = 0)
edit_menu = tk.Menu(menu_bar, tearoff = 0)
view_menu = tk.Menu(menu_bar, tearoff = 0)
option_menu = tk.Menu(menu_bar, tearoff = 0)
menu_bar.add_cascade(label = "File", menu = file_menu)
menu_bar.add_cascade(label = "Edit", menu = edit_menu)
menu_bar.add_cascade(label = "View", menu = view_menu)
menu_bar.add_cascade(label = "Option", menu = option_menu)

# Adding command for each major menu
  # File menu
file_menu.add_cascade(label = 'Something')
file_menu.add_separator()
file_menu.add_cascade(label = 'Exit')
  # Edit menu
edit_menu.add_cascade(label = 'Something')
  # View menu
view_menu.add_cascade(label = 'Something')
  # Option menu
option_menu.add_cascade(label = 'Something')

# Frame management
left_frame = tk.Frame(app, bg = "lightblue",padx = 25, pady = 8)
left_frame.grid(row=0, column=0, sticky="n")
right_frame = tk.Frame(app, bg = "red")
right_frame.grid(row=0, column=1, sticky="n")
nav_sheet_frame = tk.Frame(right_frame, bg = "blue")
nav_sheet_frame.grid(row=1, sticky="sw")

# Adding import button
import_product_button = tk.Button(left_frame, text = "Import Product",width = 50, height = 2, command = controller.importProductExcel)
import_sale_button = tk.Button(left_frame, text = "Import Sale",width = 50, height = 2, state = "disabled", command = controller.importSaleExcel)

import_product_button.grid(row=0, column=0, pady=(0,8) )
import_sale_button.grid(row=1, column=0, pady=(0,8) )

analyze_button = tk.Button(left_frame, text = "Analyze",width = 50, height = 2)
analyze_button.grid(row=2, column=0)